(() => {
  const $ = (id) => document.getElementById(id);

  const els = {
    bootScreen: $("bootScreen"),
    authScreen: $("authScreen"),
    appScreen: $("appScreen"),
    authMessage: $("authMessage"),
    resetPasswordBox: $("resetPasswordBox"),
    loginForm: $("loginForm"),
    registerForm: $("registerForm"),
    forgotPasswordBtn: $("forgotPasswordBtn"),
    updatePasswordBtn: $("updatePasswordBtn"),
    logoutBtn: $("logoutBtn"),
    mobileMenuBtn: $("mobileMenuBtn"),
    sidebar: $("sidebar"),

    orgNameLabel: $("orgNameLabel"),
    mobileOrgName: $("mobileOrgName"),
    userWelcome: $("userWelcome"),
    pageTitle: $("pageTitle"),
    pageSubtitle: $("pageSubtitle"),

    searchInput: $("searchInput"),
    ownerFilterDropdown: $("ownerFilterDropdown"),
    ownerFilterBtn: $("ownerFilterBtn"),
    ownerFilterMenu: $("ownerFilterMenu"),
    ownerFilterLabel: $("ownerFilterLabel"),
    ownerFilter: $("ownerFilter"),
    monthFilterDropdown: $("monthFilterDropdown"),
    monthFilterBtn: $("monthFilterBtn"),
    monthFilterMenu: $("monthFilterMenu"),
    monthFilterLabel: $("monthFilterLabel"),
    monthFilter: $("monthFilter"),
    selectAllLeads: $("selectAllLeads"),
    deleteSelectedBtn: $("deleteSelectedBtn"),

    importCsvBtn: $("importCsvBtn"),
    exportCsvBtn: $("exportCsvBtn"),
    csvFileInput: $("csvFileInput"),
    addLeadBtn: $("addLeadBtn"),
    addStageBtn: $("addStageBtn"),
    historyBtn: $("historyBtn"),

    pipelineScrollArea: $("pipelineScrollArea"),
    pipelineScrollTop: $("pipelineScrollTop"),
    pipelineScrollTopInner: $("pipelineScrollTopInner"),
    pipelineScrollBottom: $("pipelineScrollBottom"),
    pipelineScrollBottomInner: $("pipelineScrollBottomInner"),
    pipeline: $("pipeline"),
    leadsTableBody: $("leadsTableBody"),
    totalLeads: $("totalLeads"),
    totalValue: $("totalValue"),
    closedDeals: $("closedDeals"),
    conversionRate: $("conversionRate"),
    avgTicket: $("avgTicket"),
    organicLeads: $("organicLeads"),
    topStage: $("topStage"),
    paidRate: $("paidRate"),

    reportClosedValue: $("reportClosedValue"),
    reportPaidCount: $("reportPaidCount"),
    reportOrganicCount: $("reportOrganicCount"),
    reportAvgTicket: $("reportAvgTicket"),
    reportTopOwner: $("reportTopOwner"),
    reportTopStage: $("reportTopStage"),
    reportBestMonth: $("reportBestMonth"),
    reportClosedPlans: $("reportClosedPlans"),
    planSummaryBody: $("planSummaryBody"),

    teamList: $("teamList"),
    stagesConfigList: $("stagesConfigList"),

    modalOverlay: $("modalOverlay"),
    closeModalBtn: $("closeModalBtn"),
    cancelBtn: $("cancelBtn"),
    leadForm: $("leadForm"),
    modalTitle: $("modalTitle"),
    leadId: $("leadId"),
    name: $("name"),
    contact: $("contact"),
    owner: $("owner"),
    value: $("value"),
    startDate: $("startDate"),
    stage: $("stage"),
    socialSource: $("socialSource"),
    trafficType: $("trafficType"),
    addPlanBtn: $("addPlanBtn"),
    plansList: $("plansList"),
    addObservationBtn: $("addObservationBtn"),
    observationsList: $("observationsList"),

    stageModalOverlay: $("stageModalOverlay"),
    closeStageModalBtn: $("closeStageModalBtn"),
    cancelStageBtn: $("cancelStageBtn"),
    stageForm: $("stageForm"),
    stageModalTitle: $("stageModalTitle"),
    saveStageBtn: $("saveStageBtn"),
    stageId: $("stageId"),
    stageName: $("stageName"),
    stageColor: $("stageColor"),
    stageColorPreview: $("stageColorPreview"),
    stageType: $("stageType"),
    customStageTypeGroup: $("customStageTypeGroup"),
    customStageType: $("customStageType"),
    removeCustomTypeBtn: $("removeCustomTypeBtn"),
    savedStageTypesGroup: $("savedStageTypesGroup"),
    savedStageTypes: $("savedStageTypes"),

    historyModalOverlay: $("historyModalOverlay"),
    closeHistoryModalBtn: $("closeHistoryModalBtn"),
    refreshHistoryBtn: $("refreshHistoryBtn"),
    historyText: $("historyText")
  };

  const state = {
    supabase: null,
    currentUser: null,
    profile: null,
    stages: [],
    customStageTypes: [],
    leads: [],
    profiles: [],
    history: [],
    activeView: "funil",
    historyLoaded: false,
    profilesLoaded: false,
    chartLoader: null,
    selectedLeadIds: new Set(),
    modalPlans: [],
    modalObservations: [],
    charts: {
      pipeline: null,
      traffic: null,
      owner: null,
      monthly: null,
      social: null,
      planCount: null,
      planRevenue: null
    }
  };

  function createClient() {
    const cfg = window.APP_CONFIG || {};
    if (!cfg.supabaseUrl || !cfg.supabaseAnonKey) {
      throw new Error("Configuração do Supabase não encontrada em config.js");
    }
    state.supabase = window.supabase.createClient(cfg.supabaseUrl, cfg.supabaseAnonKey);
  }

  function setMessage(el, text, isError = false) {
    el.textContent = text || "";
    el.style.color = isError ? "#fecaca" : "#cdecd6";
  }

  function showScreen(id) {
    [els.authScreen, els.appScreen].forEach((screen) => screen.classList.add("hidden"));
    $(id).classList.remove("hidden");
    els.bootScreen.classList.add("hidden");
  }

  function brMoney(value) {
    return Number(value || 0).toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
  }

  function isMissingRelationError(error) {
    const code = String(error?.code || "").toUpperCase();
    const message = String(error?.message || "").toLowerCase();
    return code === "PGRST205" || code === "42P01" || message.includes("does not exist") || message.includes("could not find the table");
  }

  function getStoredCustomStageTypes() {
    try {
      const raw = window.localStorage.getItem("crmPax.customStageTypes");
      const parsed = JSON.parse(raw || "[]");
      return Array.isArray(parsed) ? parsed : [];
    } catch (_error) {
      return [];
    }
  }

  function persistCustomStageTypes(values) {
    const normalized = sortStageTypeNames(values);
    try {
      window.localStorage.setItem("crmPax.customStageTypes", JSON.stringify(normalized));
    } catch (_error) {
      // Ignore storage failures and keep the in-memory list.
    }
    return normalized;
  }

  function loadExternalScript(src) {
    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = src;
      script.async = true;
      script.onload = resolve;
      script.onerror = () => reject(new Error(`Falha ao carregar ${src}`));
      document.head.appendChild(script);
    });
  }

  async function ensureChartLibrary() {
    if (typeof window.Chart !== "undefined") return;
    if (!state.chartLoader) {
      state.chartLoader = loadExternalScript("https://cdn.jsdelivr.net/npm/chart.js/dist/chart.umd.min.js")
        .catch((error) => {
          state.chartLoader = null;
          throw error;
        });
    }
    await state.chartLoader;
  }

  function formatDate(value) {
    if (!value) return "-";
    const p = String(value).split("-");
    if (p.length !== 3) return value;
    return `${p[2]}/${p[1]}/${p[0]}`;
  }

  function formatMonthLabel(value) {
    if (!value || !/^\d{4}-\d{2}$/.test(String(value))) return value || "Todos os meses";
    return `${String(value).slice(5, 7)}/${String(value).slice(0, 4)}`;
  }

  function normalizeDateInput(value) {
    const raw = String(value || "").trim();
    if (!raw) return "";

    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

    const slashIso = raw.match(/^(\d{4})[\/](\d{2})[\/](\d{2})$/);
    if (slashIso) return `${slashIso[1]}-${slashIso[2]}-${slashIso[3]}`;

    const br = raw.match(/^(\d{2})[\/-](\d{2})[\/-](\d{4})$/);
    if (br) return `${br[3]}-${br[2]}-${br[1]}`;

    const parsed = new Date(raw);
    if (!Number.isNaN(parsed.getTime())) {
      const year = parsed.getFullYear();
      const month = String(parsed.getMonth() + 1).padStart(2, "0");
      const day = String(parsed.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }

    return "";
  }

  function cleanObservationList(observations) {
    return (Array.isArray(observations) ? observations : [])
      .map((item) => ({
        date: normalizeDateInput(item?.date || ""),
        text: String(item?.text || "").trim()
      }))
      .filter((item) => item.text);
  }

  function cleanPlanList(plans) {
    return (Array.isArray(plans) ? plans : [])
      .map((item) => {
        const name = String(item?.name || "").trim();
        const rawValue = item?.value;

        let value = 0;
        if (rawValue !== null && rawValue !== undefined && rawValue !== "") {
          const normalized = String(rawValue)
            .replace(/r\$\s*/gi, "")
            .replace(/\./g, "")
            .replace(",", ".")
            .replace(/[^0-9.-]/g, "");
          value = Number(normalized || 0);
        }

        return {
          name,
          value: Number.isFinite(value) ? value : 0
        };
      })
      .filter((item) => item.name);
  }


  function getLeadMeta(rawNotes, leadValue = 0) {
    const raw = String(rawNotes || "");
    const prefix = "__CRM_META__";

    if (!raw.startsWith(prefix)) {
      const text = raw.trim();
      return {
        plan: "",
        plans: [],
        legacyText: text,
        observations: text ? [{ date: "", text }] : []
      };
    }

    try {
      const parsed = JSON.parse(raw.slice(prefix.length));
      const legacyText = String(parsed?.legacyText || "").trim();
      const observations = cleanObservationList(parsed?.observations || []);
      const plans = cleanPlanList(parsed?.plans || []);
      const legacyPlan = String(parsed?.plan || "").trim();
      if (!plans.length && legacyPlan) {
        plans.push({ name: legacyPlan, value: Number.isFinite(Number(leadValue)) ? Number(leadValue) : 0 });
      }
      if (!observations.length && legacyText) observations.push({ date: "", text: legacyText });
      return {
        plan: legacyPlan || (plans[0]?.name || ""),
        plans,
        legacyText,
        observations
      };
    } catch (_error) {
      const text = raw.trim();
      return {
        plan: "",
        plans: [],
        legacyText: text,
        observations: text ? [{ date: "", text }] : []
      };
    }
  }

  function serializeLeadMeta(meta = {}) {
    const legacyText = String(meta.legacyText || "").trim();
    const plans = cleanPlanList(meta.plans || []);
    const observations = cleanObservationList(meta.observations || []);
    const plan = plans[0]?.name || String(meta.plan || "").trim();

    if (!plan && !plans.length && !observations.length) return legacyText;

    return `__CRM_META__${JSON.stringify({ plan, plans, legacyText, observations })}`;
  }

  function normalizeLead(lead) {
    return {
      ...lead,
      start_date: normalizeDateInput(lead?.start_date || "") || String(lead?.start_date || "").trim(),
      _meta: getLeadMeta(lead?.notes || "", lead?.value || 0)
    };
  }

  function getLeadPlans(lead) {
    return cleanPlanList(lead?._meta?.plans || []);
  }

  function getLeadPlan(lead) {
    const plans = getLeadPlans(lead);
    if (plans.length) return plans.map((item) => item.name).join(", ");
    return String(lead?._meta?.plan || "").trim();
  }

  function getLeadObservations(lead) {
    return cleanObservationList(lead?._meta?.observations || []);
  }

  function getLeadLatestObservation(lead) {
    const observations = getLeadObservations(lead);
    return observations[observations.length - 1] || null;
  }

  function getLeadSearchText(lead) {
    return [
      lead?.name,
      lead?.contact,
      lead?.owner,
      lead?.social_source,
      lead?._meta?.legacyText,
      getLeadPlan(lead),
      ...getLeadPlans(lead).map((item) => `${item.name} ${item.value}`),
      ...getLeadObservations(lead).map((item) => item.text)
    ].join(" ").toLowerCase();
  }

  function getLeadMonthKey(lead) {
    const fromStart = normalizeDateInput(lead?.start_date || "");
    if (fromStart) return fromStart.slice(0, 7);

    const created = normalizeDateInput(String(lead?.created_at || "").slice(0, 10));
    return created ? created.slice(0, 7) : "";
  }

  function hasLeadValue(lead) {
    return Number(lead?.value || 0) > 0;
  }

  function getClosedStageIds() {
    return state.stages.filter((stage) => stage.stage_type === "fechado").map((stage) => stage.id);
  }

  function getQualifiedClosedLeads(leads) {
    const closedIds = getClosedStageIds();
    return (Array.isArray(leads) ? leads : []).filter((lead) => closedIds.includes(lead.stage_id) && hasLeadValue(lead));
  }

  function syncSelectedLeadIds() {
    const validIds = new Set(state.leads.map((lead) => lead.id));
    state.selectedLeadIds = new Set(
      [...state.selectedLeadIds].filter((id) => validIds.has(id))
    );
  }

  function closeFilterDropdowns(except = null) {
    [
      { dropdown: els.ownerFilterDropdown, btn: els.ownerFilterBtn, menu: els.ownerFilterMenu },
      { dropdown: els.monthFilterDropdown, btn: els.monthFilterBtn, menu: els.monthFilterMenu }
    ].forEach((item) => {
      if (!item.dropdown || item.dropdown === except) return;
      item.dropdown.classList.remove("open");
      item.menu?.classList.add("hidden");
      item.btn?.setAttribute("aria-expanded", "false");
    });
  }

  function setFilterDropdownOpen(dropdown, btn, menu, shouldOpen) {
    if (!dropdown || !btn || !menu) return;
    dropdown.classList.toggle("open", shouldOpen);
    menu.classList.toggle("hidden", !shouldOpen);
    btn.setAttribute("aria-expanded", shouldOpen ? "true" : "false");
  }

  function syncFilterButtonLabel(selectEl, labelEl, defaultLabel, formatLabel = (value, text) => text || value) {
    if (!selectEl || !labelEl) return;
    const selectedOption = selectEl.options[selectEl.selectedIndex];
    const value = selectEl.value;
    labelEl.textContent = value ? formatLabel(value, selectedOption?.textContent || value) : defaultLabel;
  }

  function renderFilterOptions(selectEl, menuEl, labelEl, defaultLabel, formatLabel = (value, text) => text || value) {
    if (!selectEl || !menuEl) return;

    const options = [...selectEl.options].map((option) => ({
      value: option.value,
      label: option.textContent || ""
    }));

    menuEl.innerHTML = options.map((option) => `
      <button
        type="button"
        class="filter-option ${option.value === selectEl.value ? "active" : ""}"
        data-value="${escapeHtml(option.value)}"
      >
        ${escapeHtml(option.label)}
      </button>
    `).join("");

    syncFilterButtonLabel(selectEl, labelEl, defaultLabel, formatLabel);

    menuEl.querySelectorAll(".filter-option").forEach((button) => {
      button.addEventListener("click", () => {
        selectEl.value = button.dataset.value || "";
        syncFilterButtonLabel(selectEl, labelEl, defaultLabel, formatLabel);
        closeFilterDropdowns();
        renderAll();
      });
    });
  }

  function stageTypeLabel(type, customStageType = "") {
    const map = {
      andamento: "Andamento",
      fechado: "Fechado",
      cancelado: "Cancelado",
      espera: "Espera",
      personalizado: customStageType || "Personalizado"
    };
    return map[type] || customStageType || "Andamento";
  }

  function sortStageTypeNames(values) {
    return [...new Set(
      (Array.isArray(values) ? values : [])
        .map((item) => String(item || "").trim())
        .filter(Boolean)
    )].sort((a, b) => a.localeCompare(b, "pt-BR"));
  }

  function getCustomStageTypes() {
    return sortStageTypeNames([
      ...state.customStageTypes,
      ...getStoredCustomStageTypes(),
      ...state.stages.map((stage) => String(stage.custom_stage_type || "").trim())
    ]);
  }

  async function saveCustomStageType(name) {
    const normalized = String(name || "").trim();
    if (!normalized) return true;

    state.customStageTypes = persistCustomStageTypes([...state.customStageTypes, normalized]);

    const { error } = await state.supabase
      .from("stage_type_catalog")
      .upsert(
        { name: normalized, created_by: state.currentUser?.id || null },
        { onConflict: "name" }
      );

    if (error && !isMissingRelationError(error)) {
      console.error("Erro ao salvar tipo personalizado:", error);
      return false;
    }

    return true;
  }

  async function deleteCustomStageTypeFromCatalog(name) {
    const normalized = String(name || "").trim();
    if (!normalized) return true;

    state.customStageTypes = persistCustomStageTypes(
      state.customStageTypes.filter((item) => item !== normalized)
    );

    const { error } = await state.supabase
      .from("stage_type_catalog")
      .delete()
      .eq("name", normalized);

    if (error && !isMissingRelationError(error)) {
      console.error("Erro ao excluir tipo personalizado:", error);
      return false;
    }

    return true;
  }

  function refreshStageTypeOptions(selectedValue = "", customValue = "") {
    if (!els.stageType) return;

    const baseOptions = [
      ["andamento", "Andamento"],
      ["fechado", "Fechado"],
      ["cancelado", "Cancelado"],
      ["espera", "Espera"]
    ];

    const customOptions = getCustomStageTypes().map((item) => [`custom:${item}`, item]);
    const finalOptions = [...baseOptions, ...customOptions, ["personalizado", "+ Novo tipo personalizado"]];

    els.stageType.innerHTML = finalOptions.map(([value, label]) => `<option value="${escapeHtml(value)}">${escapeHtml(label)}</option>`).join("");

    const resolvedValue = selectedValue || (customValue ? `custom:${customValue}` : "andamento");
    els.stageType.value = finalOptions.some(([value]) => value === resolvedValue) ? resolvedValue : (customValue ? "personalizado" : "andamento");
    if (els.customStageType) els.customStageType.value = els.stageType.value === "personalizado" ? (customValue || "") : "";
    renderCustomStageTypeList();
    toggleCustomStageTypeField();
  }

  function renderCustomStageTypeList() {
    if (!els.savedStageTypes || !els.savedStageTypesGroup) return;

    const types = getCustomStageTypes();
    els.savedStageTypesGroup.classList.remove("hidden");

    if (!types.length) {
      els.savedStageTypes.innerHTML = '<div class="saved-stage-types-empty">Nenhum tipo personalizado salvo ainda.</div>';
      return;
    }

    const selected = String(els.stageType?.value || "");
    els.savedStageTypes.innerHTML = types.map((type) => {
      const active = selected === `custom:${type}`;
      return `
        <div class="saved-stage-type ${active ? "active" : ""}">
          <button type="button" class="saved-stage-type-select" data-stage-type="${escapeHtml(type)}">${escapeHtml(type)}</button>
          <button type="button" class="saved-stage-type-delete" data-stage-type-delete="${escapeHtml(type)}" aria-label="Excluir tipo ${escapeHtml(type)}">&times;</button>
        </div>
      `;
    }).join("");
  }

  function toggleCustomStageTypeField() {
    if (!els.customStageTypeGroup) return;
    const selected = String(els.stageType?.value || "andamento");
    const isNewCustom = selected === "personalizado";
    const isExistingCustom = selected.startsWith("custom:");
    els.customStageTypeGroup.classList.toggle("hidden", !(isNewCustom || isExistingCustom));
    if (els.customStageType) {
      els.customStageType.disabled = !isNewCustom;
      els.customStageType.value = isNewCustom ? els.customStageType.value : (isExistingCustom ? selected.replace(/^custom:/, "") : "");
    }
    if (els.removeCustomTypeBtn) els.removeCustomTypeBtn.classList.toggle("hidden", !(isNewCustom || isExistingCustom));
    renderCustomStageTypeList();
  }

  async function removeCurrentCustomStageType() {
    const selected = String(els.stageType?.value || "");
    if (!selected.startsWith("custom:")) {
      els.stageType.value = "andamento";
      if (els.customStageType) els.customStageType.value = "";
      toggleCustomStageTypeField();
      return;
    }

    const customValue = selected.replace(/^custom:/, "");
    if (!customValue) return;
    if (!confirm(`Remover o tipo personalizado "${customValue}" da lista? Os pipelines com esse tipo voltarão para Andamento.`)) return;

    const affected = state.stages.filter((stage) => String(stage.custom_stage_type || "").trim() === customValue);
    if (affected.length) {
      const { error } = await state.supabase
        .from("stages")
        .update({ stage_type: "andamento", custom_stage_type: null })
        .in("id", affected.map((stage) => stage.id));
      if (error) return alert(`Erro no Supabase: ${error.message}`);
    }

    const removed = await deleteCustomStageTypeFromCatalog(customValue);
    if (!removed) return alert("Não foi possível excluir o tipo personalizado da lista.");

    els.stageType.value = "andamento";
    if (els.customStageType) els.customStageType.value = "";
    toggleCustomStageTypeField();
    await loadAppData({ includeProfiles: state.profilesLoaded });
  }


  function renderPlanItems() {
    if (!els.plansList) return;

    if (!state.modalPlans.length) {
      els.plansList.innerHTML = '<div class="plan-empty">Nenhum plano adicionado ainda.</div>';
      return;
    }

    els.plansList.innerHTML = state.modalPlans.map((item, index) => `
      <div class="plan-item" data-index="${index}">
        <div class="plan-item-grid">
          <input type="text" class="plan-name-input" placeholder="Nome do plano" value="${escapeHtml(item.name || "")}" />
          <input type="number" class="plan-value-input" placeholder="Valor do plano" min="0" step="0.01" value="${escapeHtml(item.value ?? "")}" />
        </div>
        <div class="plan-item-actions">
          <button type="button" class="btn btn-danger plan-remove-btn">Excluir</button>
        </div>
      </div>
    `).join("");
  }

  function addPlanFromDraft() {
    state.modalPlans.push({ name: "", value: "" });
    renderPlanItems();
  }

  function setupPlanListEvents() {
    els.plansList?.addEventListener("input", (event) => {
      const item = event.target.closest(".plan-item");
      if (!item) return;
      const index = Number(item.dataset.index);
      if (Number.isNaN(index) || !state.modalPlans[index]) return;
      if (event.target.classList.contains("plan-name-input")) {
        state.modalPlans[index].name = event.target.value;
      }
      if (event.target.classList.contains("plan-value-input")) {
        state.modalPlans[index].value = event.target.value;
      }
    });

    els.plansList?.addEventListener("click", (event) => {
      const button = event.target.closest(".plan-remove-btn");
      if (!button) return;
      const item = button.closest(".plan-item");
      const index = Number(item?.dataset.index);
      if (Number.isNaN(index)) return;
      state.modalPlans.splice(index, 1);
      renderPlanItems();
    });
  }

  function renderObservationItems() {
    if (!els.observationsList) return;

    if (!state.modalObservations.length) {
      els.observationsList.innerHTML = '<div class="observation-empty">Nenhuma observação adicionada ainda.</div>';
      return;
    }

    els.observationsList.innerHTML = state.modalObservations.map((item, index) => `
      <div class="observation-item" data-index="${index}">
        <div class="observation-item-header">
          <input type="date" class="observation-date-input" value="${escapeHtml(item.date || "")}" />
        </div>
        <textarea class="observation-text-input" rows="4" placeholder="Digite a observação">${escapeHtml(item.text || "")}</textarea>
        <div class="observation-item-actions">
          <button type="button" class="btn btn-danger observation-remove-btn">Excluir</button>
        </div>
      </div>
    `).join("");
  }

  function addObservationFromDraft() {
    state.modalObservations.push({ date: "", text: "" });
    renderObservationItems();
  }

  function setupObservationListEvents() {
    els.observationsList?.addEventListener("input", (event) => {
      const item = event.target.closest(".observation-item");
      if (!item) return;
      const index = Number(item.dataset.index);
      if (Number.isNaN(index) || !state.modalObservations[index]) return;
      if (event.target.classList.contains("observation-date-input")) {
        state.modalObservations[index].date = normalizeDateInput(event.target.value);
      }
      if (event.target.classList.contains("observation-text-input")) {
        state.modalObservations[index].text = event.target.value;
      }
    });

    els.observationsList?.addEventListener("click", (event) => {
      const button = event.target.closest(".observation-remove-btn");
      if (!button) return;
      const item = button.closest(".observation-item");
      const index = Number(item?.dataset.index);
      if (Number.isNaN(index)) return;
      state.modalObservations.splice(index, 1);
      renderObservationItems();
    });
  }

  function syncPipelineScrollBars(source = null) {
    const area = els.pipelineScrollArea;
    const top = els.pipelineScrollTop;
    const bottom = els.pipelineScrollBottom;
    const width = area?.scrollWidth || 0;
    const client = area?.clientWidth || 0;

    if (els.pipelineScrollTopInner) els.pipelineScrollTopInner.style.width = `${width}px`;
    if (els.pipelineScrollBottomInner) els.pipelineScrollBottomInner.style.width = `${width}px`;

    [top, bottom].forEach((bar) => {
      if (!bar) return;
      bar.classList.toggle("hidden", width <= client);
    });

    if (!area) return;

    const left = source?.scrollLeft ?? area.scrollLeft;
    if (source !== area) area.scrollLeft = left;
    if (top && source !== top) top.scrollLeft = left;
    if (bottom && source !== bottom) bottom.scrollLeft = left;
  }

  function attachPipelineScrollEvents() {
    const area = els.pipelineScrollArea;
    const top = els.pipelineScrollTop;
    const bottom = els.pipelineScrollBottom;
    if (!area) return;

    area.addEventListener("scroll", () => syncPipelineScrollBars(area));
    top?.addEventListener("scroll", () => syncPipelineScrollBars(top));
    bottom?.addEventListener("scroll", () => syncPipelineScrollBars(bottom));
    window.addEventListener("resize", () => syncPipelineScrollBars());
  }

  function escapeHtml(value) {
    return String(value || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function getUserDisplayName() {
    return (
      state.profile?.full_name ||
      state.currentUser?.user_metadata?.full_name ||
      state.currentUser?.email ||
      "Usuário"
    );
  }

  function updateStageColorPreview(value) {
    const color = String(value || "#1f9d55").toUpperCase();
    if (els.stageColorPreview) {
      els.stageColorPreview.textContent = color;
      els.stageColorPreview.style.borderColor = `${color}55`;
      els.stageColorPreview.style.background = `${color}22`;
      els.stageColorPreview.style.color = color;
    }
  }

  function getStageById(id) {
    return state.stages.find((stage) => stage.id === id) || null;
  }

  async function persistStagePositions(stages) {
    const updates = stages.map((stage, index) =>
      state.supabase.from("stages").update({ position: index }).eq("id", stage.id)
    );

    const results = await Promise.all(updates);
    const failed = results.find((result) => result.error);
    if (failed?.error) {
      throw failed.error;
    }
  }

  async function moveStagePosition(stageId, direction) {
    const currentIndex = state.stages.findIndex((stage) => stage.id === stageId);
    if (currentIndex === -1) return;

    const targetIndex = direction === "left" ? currentIndex - 1 : currentIndex + 1;
    if (targetIndex < 0 || targetIndex >= state.stages.length) return;

    const reordered = [...state.stages];
    const [stage] = reordered.splice(currentIndex, 1);
    reordered.splice(targetIndex, 0, stage);

    try {
      await persistStagePositions(reordered);
      await logChange(
        "reorder",
        "stage",
        stageId,
        `A ordem do funil foi alterada por ${getUserDisplayName()}.`,
        {
          stage_id: stageId,
          stage_name: stage.name,
          from_position: currentIndex,
          to_position: targetIndex
        }
      );
      await loadAppData({ includeProfiles: state.profilesLoaded });
    } catch (error) {
      alert(error.message || "Não foi possível reordenar o pipeline.");
    }
  }

  async function bootstrap() {
    const params = new URLSearchParams(window.location.hash.replace(/^#/, ""));

    if (params.get("type") === "recovery") {
      els.resetPasswordBox.classList.remove("hidden");
      showScreen("authScreen");
      setMessage(els.authMessage, "Digite sua nova senha e salve.");
      return;
    }

    const { data, error } = await state.supabase.auth.getSession();
    if (error) console.error(error);

    state.currentUser = data?.session?.user || null;

    if (state.currentUser) {
      await ensureProfile();
      await enterApp();
      return;
    }

    showScreen("authScreen");
  }

  async function ensureProfile() {
    if (!state.currentUser) return;

    const defaultName =
      state.currentUser.user_metadata?.full_name ||
      state.currentUser.email?.split("@")[0] ||
      "Usuário";

    const payload = {
      id: state.currentUser.id,
      full_name: defaultName,
      email: state.currentUser.email
    };

    const { data: profile, error: upsertError } = await state.supabase
      .from("profiles")
      .upsert(payload, { onConflict: "id" })
      .select("*")
      .maybeSingle();

    if (upsertError) {
      console.error("Erro ao garantir profile:", upsertError);
    }

    state.profile = profile || payload;
  }

  async function enterApp() {
    showScreen("appScreen");
    els.orgNameLabel.textContent = "CRM Pax";
    els.mobileOrgName.textContent = "CRM Pax";
    els.userWelcome.textContent = getUserDisplayName();

    await loadAppData({ includeProfiles: state.profilesLoaded });
    bindView("funil");
  }

  async function loadAppData(options = {}) {
    const includeProfiles = options.includeProfiles === true;
    const [stagesRes, leadsRes, profilesRes, stageTypesRes] = await Promise.all([
      state.supabase.from("stages").select("*").order("position", { ascending: true }),
      state.supabase.from("leads").select("*").order("created_at", { ascending: false }),
      includeProfiles
        ? state.supabase.from("profiles").select("*").order("full_name", { ascending: true })
        : Promise.resolve({ data: state.profiles, error: null }),
      state.supabase.from("stage_type_catalog").select("name").order("name", { ascending: true })
    ]);

    if (stagesRes.error) console.error(stagesRes.error);
    if (leadsRes.error) console.error(leadsRes.error);
    if (includeProfiles && profilesRes.error) console.error(profilesRes.error);
    if (stageTypesRes.error && !isMissingRelationError(stageTypesRes.error)) console.error(stageTypesRes.error);

    const firstError =
      stagesRes.error ||
      leadsRes.error ||
      profilesRes.error ||
      (stageTypesRes.error && !isMissingRelationError(stageTypesRes.error) ? stageTypesRes.error : null);
    if (firstError) {
      alert(`Erro ao carregar dados do Supabase: ${firstError.message}`);
      return;
    }

    state.stages = stagesRes.data || [];
    state.leads = (leadsRes.data || []).map(normalizeLead);
    state.customStageTypes = persistCustomStageTypes([
      ...getStoredCustomStageTypes(),
      ...(stageTypesRes.data || []).map((item) => item?.name)
    ]);
    if (includeProfiles) {
      state.profiles = profilesRes.data || [];
      state.profilesLoaded = true;
    }
    syncSelectedLeadIds();

    if (state.stages.length === 0) {
      await seedDefaultStages();
      return loadAppData({ includeProfiles });
    }

    renderAll();
  }

  async function loadProfilesIfNeeded(force = false) {
    if (!force && state.profilesLoaded) return;

    const { data, error } = await state.supabase
      .from("profiles")
      .select("*")
      .order("full_name", { ascending: true });

    if (error) {
      console.error(error);
      return;
    }

    state.profiles = data || [];
    state.profilesLoaded = true;
    if (state.activeView === "equipe") renderTeam();
    if (!els.historyModalOverlay.classList.contains("hidden")) renderHistoryText();
  }

  async function loadHistory(force = false) {
    if (!force && state.historyLoaded) return;

    const { data, error } = await state.supabase
      .from("change_history")
      .select("*")
      .order("created_at", { ascending: false })
      .limit(300);

    if (error) {
      console.error(error);
      if (force) alert(`Erro ao carregar histórico: ${error.message}`);
      return;
    }

    state.history = data || [];
    state.historyLoaded = true;
    renderHistoryText();
  }
  
  async function seedDefaultStages() {
    const defaults = [
      { name: "Novos Leads", color: "#3b82f6", stage_type: "andamento", position: 0, created_by: state.currentUser.id },
      { name: "Em Contato", color: "#f59e0b", stage_type: "andamento", position: 1, created_by: state.currentUser.id },
      { name: "Proposta", color: "#8b5cf6", stage_type: "andamento", position: 2, created_by: state.currentUser.id },
      { name: "Fechado", color: "#22c55e", stage_type: "fechado", position: 3, created_by: state.currentUser.id },
      { name: "Cancelado", color: "#ef4444", stage_type: "cancelado", position: 4, created_by: state.currentUser.id }
    ];

    const { error } = await state.supabase.from("stages").insert(defaults);
    if (error) alert(error.message);
  }

  function getStageName(stageId) {
    return state.stages.find((s) => s.id === stageId)?.name || "-";
  }

  function getProfileNameById(userId) {
    return state.profiles.find((p) => p.id === userId)?.full_name || "-";
  }

  function getFilteredLeads(options = {}) {
    const search = els.searchInput.value.trim().toLowerCase();
    const owner = els.ownerFilter.value;
    const month = options.ignoreMonth ? "" : els.monthFilter.value;

    return state.leads.filter((lead) => {
      const matchesSearch = !search || getLeadSearchText(lead).includes(search);
      const matchesOwner = !owner || lead.owner === owner;
      const matchesMonth = !month || getLeadMonthKey(lead) === month;

      return matchesSearch && matchesOwner && matchesMonth;
    });
  }

  function populateFilters() {
    const owners = [...new Set(state.leads.map((x) => x.owner).filter(Boolean))].sort();
    const months = [...new Set(state.leads.map((lead) => getLeadMonthKey(lead)).filter(Boolean))].sort();

    const currentOwner = els.ownerFilter.value;
    const currentMonth = els.monthFilter.value;

    els.ownerFilter.innerHTML =
      '<option value="">Todos os responsáveis</option>' +
      owners.map((o) => `<option value="${escapeHtml(o)}">${escapeHtml(o)}</option>`).join("");

    els.monthFilter.innerHTML =
      '<option value="">Todos os meses</option>' +
      months.map((m) => `<option value="${m}">${formatMonthLabel(m)}</option>`).join("");

    els.stage.innerHTML = state.stages
      .map((s) => `<option value="${s.id}">${escapeHtml(s.name)}</option>`)
      .join("");

    refreshStageTypeOptions(els.stageType?.value || "andamento");

    els.ownerFilter.value = owners.includes(currentOwner) ? currentOwner : "";
    els.monthFilter.value = months.includes(currentMonth) ? currentMonth : "";

    renderFilterOptions(
      els.ownerFilter,
      els.ownerFilterMenu,
      els.ownerFilterLabel,
      "Todos os responsáveis",
      (_value, text) => text
    );

    renderFilterOptions(
      els.monthFilter,
      els.monthFilterMenu,
      els.monthFilterLabel,
      "Todos os meses",
      (value) => formatMonthLabel(value)
    );
  }

  function getDashboardMetrics(filtered = getFilteredLeads()) {
    const closedIds = getClosedStageIds();
    const total = filtered.length;
    const qualifiedClosed = getQualifiedClosedLeads(filtered);
    const closed = qualifiedClosed.length;
    const conversion = total ? ((closed / total) * 100) : 0;
    const totalValue = qualifiedClosed.reduce((sum, item) => sum + Number(item.value || 0), 0);
    const avgTicket = qualifiedClosed.length ? totalValue / qualifiedClosed.length : 0;
    const paidCount = filtered.filter((lead) => String(lead.traffic_type || "").toLowerCase() === "pago").length;
    const organicCount = filtered.filter((lead) => String(lead.traffic_type || "").toLowerCase() === "organico").length;

    const byStage = state.stages.map((stage) => ({
      id: stage.id,
      name: stage.name,
      color: stage.color,
      count: filtered.filter((lead) => lead.stage_id === stage.id).length
    }));
    const topStage = [...byStage].sort((a, b) => b.count - a.count)[0];

    const ownerTotals = qualifiedClosed.reduce((acc, lead) => {
      const key = lead.owner || "Sem responsável";
      acc[key] = (acc[key] || 0) + Number(lead.value || 0);
      return acc;
    }, {});
    const topOwner = Object.entries(ownerTotals).sort((a, b) => b[1] - a[1])[0];

    const monthTotals = filtered.reduce((acc, lead) => {
      const key = getLeadMonthKey(lead);
      if (!key) return acc;
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});
    const bestMonth = Object.entries(monthTotals).sort((a, b) => b[1] - a[1])[0];

    const planSummaryMap = qualifiedClosed.reduce((acc, lead) => {
      const leadPlans = getLeadPlans(lead);
      const fallbackPlans = leadPlans.length ? leadPlans : (getLeadPlan(lead) ? [{ name: getLeadPlan(lead), value: Number(lead.value || 0) }] : []);
      fallbackPlans.forEach((planItem) => {
        const plan = planItem.name || "Sem plano";
        const unitValue = Number(planItem.value || 0);
        const key = `${plan}__${unitValue}`;
        if (!acc[key]) acc[key] = { plan, unitValue, count: 0, totalValue: 0 };
        acc[key].count += 1;
        acc[key].totalValue += unitValue;
      });
      return acc;
    }, {});

    const planSummary = Object.values(planSummaryMap)
      .sort((a, b) => (b.totalValue - a.totalValue) || a.plan.localeCompare(b.plan, "pt-BR"));

    return { total, totalValue, closed, conversion, avgTicket, paidCount, organicCount, byStage, topStage, topOwner, bestMonth, planSummary };
  }

  function renderStats() {
    const metrics = getDashboardMetrics();

    els.totalLeads.textContent = metrics.total;
    els.totalValue.textContent = brMoney(metrics.totalValue);
    els.closedDeals.textContent = metrics.closed;
    els.conversionRate.textContent = `${metrics.conversion.toFixed(1)}%`;
    els.avgTicket.textContent = brMoney(metrics.avgTicket);
    els.topStage.textContent = metrics.topStage?.name || "-";
    els.paidRate.textContent = metrics.paidCount;
    if (els.organicLeads) els.organicLeads.textContent = metrics.organicCount;

    els.reportClosedValue.textContent = brMoney(metrics.totalValue);
    els.reportPaidCount.textContent = metrics.paidCount;
    els.reportOrganicCount.textContent = metrics.organicCount;
    els.reportAvgTicket.textContent = brMoney(metrics.avgTicket);
    els.reportTopOwner.textContent = metrics.topOwner?.[0] || "-";
    els.reportTopStage.textContent = metrics.topStage?.name || "-";
    els.reportBestMonth.textContent = metrics.bestMonth?.[0] ? formatMonthLabel(metrics.bestMonth[0]) : "-";
    if (els.reportClosedPlans) els.reportClosedPlans.textContent = metrics.planSummary.reduce((sum, item) => sum + item.count, 0);

    if (els.planSummaryBody) {
      els.planSummaryBody.innerHTML = metrics.planSummary.length
        ? metrics.planSummary.map((item) => `
          <tr>
            <td>${escapeHtml(item.plan)}</td>
            <td>${brMoney(item.unitValue)}</td>
            <td>${item.count}</td>
            <td>${brMoney(item.totalValue)}</td>
          </tr>
        `).join("")
        : '<tr><td colspan="4" class="empty-state">Nenhum fechamento com valor encontrado.</td></tr>';
    }
  }

  function renderPipeline() {
    const filtered = getFilteredLeads();

    els.pipeline.innerHTML = state.stages.map((stage, index) => {
      const leads = filtered.filter((lead) => lead.stage_id === stage.id);

      const cards = leads.length
        ? leads.map((lead) => `
          <article class="card" draggable="true" data-lead-id="${lead.id}">
            <div class="card-top">
              <div>
                <div class="card-title">${escapeHtml(lead.name)}</div>
                <div class="card-value">${hasLeadValue(lead) ? brMoney(lead.value) : "Sem valor"}</div>
              </div>
              <span class="status-pill" style="background:${stage.color}22;color:${stage.color};border-color:${stage.color}55">
                ${escapeHtml(stageTypeLabel(stage.stage_type, stage.custom_stage_type))}
              </span>
            </div>

            <div class="card-meta">
              <span><strong>Contato:</strong> ${escapeHtml(lead.contact || "-")}</span>
              <span><strong>Responsável:</strong> ${escapeHtml(lead.owner || "-")}</span>
              <span><strong>Início:</strong> ${formatDate(lead.start_date)}</span>
              <span><strong>Origem:</strong> ${escapeHtml(lead.traffic_type || "-")}</span>
              <span><strong>Rede:</strong> ${escapeHtml(lead.social_source || "-")}</span>
              <span><strong>Plano:</strong> ${escapeHtml(getLeadPlan(lead) || "-")}</span>
            </div>

            ${getLeadLatestObservation(lead) ? `<div class="card-notes"><strong>Última observação:</strong> ${escapeHtml(getLeadLatestObservation(lead).text)}${getLeadLatestObservation(lead).date ? `<small>${formatDate(getLeadLatestObservation(lead).date)}</small>` : ""}</div>` : ""}

            <div class="card-actions">
              <button type="button" class="edit-btn" data-action="edit-lead" data-id="${lead.id}">Editar</button>
              <button type="button" class="delete-btn" data-action="delete-lead" data-id="${lead.id}">Excluir</button>
            </div>
          </article>
        `).join("")
        : '<div class="empty-state">Nenhum lead nesta etapa.</div>';

      return `
        <section class="column" data-stage-id="${stage.id}">
          <header class="column-header">
            <div class="column-title-wrap">
              <div class="column-title">
                <div class="column-accent" style="background:${stage.color}"></div>
                <div>
                  <h3>${escapeHtml(stage.name)}</h3>
                  <span>${escapeHtml(stageTypeLabel(stage.stage_type, stage.custom_stage_type))}</span>
                </div>
              </div>
              <div class="column-order-actions">
                <button type="button" class="icon-btn" data-stage-action="move-left" data-id="${stage.id}" ${index === 0 ? "disabled" : ""} title="Mover para a esquerda">&larr;</button>
                <button type="button" class="icon-btn" data-stage-action="move-right" data-id="${stage.id}" ${index === state.stages.length - 1 ? "disabled" : ""} title="Mover para a direita">&rarr;</button>
              </div>
            </div>
            <span class="badge">${leads.length}</span>
          </header>
          <div class="column-body">${cards}</div>
        </section>
      `;
    }).join("");

    bindPipelineEvents();
    requestAnimationFrame(syncPipelineScrollBars);
  }

  function updateBulkDeleteButton() {
    if (!els.deleteSelectedBtn) return;

    const count = state.selectedLeadIds.size;
    els.deleteSelectedBtn.textContent = count ? `Excluir selecionados (${count})` : "Excluir selecionados";
    els.deleteSelectedBtn.classList.toggle("hidden", count === 0);
  }

  function syncSelectAllCheckbox(filtered) {
    if (!els.selectAllLeads) return;

    const visibleIds = filtered.map((lead) => lead.id);

    if (!visibleIds.length) {
      els.selectAllLeads.checked = false;
      els.selectAllLeads.indeterminate = false;
      return;
    }

    const selectedVisibleCount = visibleIds.filter((id) => state.selectedLeadIds.has(id)).length;
    els.selectAllLeads.checked = selectedVisibleCount === visibleIds.length;
    els.selectAllLeads.indeterminate = selectedVisibleCount > 0 && selectedVisibleCount < visibleIds.length;
  }

  async function deleteSelectedLeads() {
    const ids = [...state.selectedLeadIds];
    if (!ids.length) return;

    const selectedLeads = state.leads.filter((lead) => ids.includes(lead.id));
    if (!confirm(`Excluir ${ids.length} lead(s) selecionado(s)?`)) return;

    const { error } = await state.supabase
      .from("leads")
      .delete()
      .in("id", ids);

    if (error) return alert(`Erro no Supabase: ${error.message}`);

    await logChange(
      "bulk_delete",
      "lead",
      null,
      `${ids.length} lead(s) foram excluídos por ${getUserDisplayName()}.`,
      {
        deleted_ids: ids,
        deleted_names: selectedLeads.map((lead) => lead.name)
      }
    );

    state.selectedLeadIds.clear();
    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  function renderLeadTable() {
    const filtered = getFilteredLeads();

    if (!filtered.length) {
      els.leadsTableBody.innerHTML = `
        <tr>
          <td colspan="10" class="empty-state">Nenhum lead encontrado.</td>
        </tr>
      `;
      updateBulkDeleteButton();
      syncSelectAllCheckbox(filtered);
      return;
    }

    els.leadsTableBody.innerHTML = filtered.map((lead) => `
      <tr>
        <td class="select-col">
          <input
            type="checkbox"
            class="lead-check"
            data-id="${lead.id}"
            ${state.selectedLeadIds.has(lead.id) ? "checked" : ""}
            aria-label="Selecionar ${escapeHtml(lead.name)}"
          />
        </td>
        <td>${escapeHtml(lead.name)}</td>
        <td>${escapeHtml(lead.contact || "-")}</td>
        <td>${escapeHtml(lead.owner || "-")}</td>
        <td>${hasLeadValue(lead) ? brMoney(lead.value) : "-"}</td>
        <td>${formatDate(lead.start_date)}</td>
        <td>${escapeHtml(lead.social_source || "-")}</td>
        <td>${escapeHtml(lead.traffic_type || "-")}</td>
        <td>${escapeHtml(getLeadPlan(lead) || "-")}</td>
        <td>${escapeHtml(getStageName(lead.stage_id))}</td>
        <td>
          <div class="table-actions">
            <button type="button" class="edit-btn" data-action="edit-lead" data-id="${lead.id}">Editar</button>
            <button type="button" class="delete-btn" data-action="delete-lead" data-id="${lead.id}">Excluir</button>
          </div>
        </td>
      </tr>
    `).join("");

    document.querySelectorAll(".lead-check").forEach((checkbox) => {
      checkbox.addEventListener("change", () => {
        const id = checkbox.dataset.id;
        if (!id) return;

        if (checkbox.checked) {
          state.selectedLeadIds.add(id);
        } else {
          state.selectedLeadIds.delete(id);
        }

        updateBulkDeleteButton();
        syncSelectAllCheckbox(filtered);
      });
    });

    updateBulkDeleteButton();
    syncSelectAllCheckbox(filtered);
  }

  function renderTeam() {
    const activeIds = new Set([
      ...state.leads.map((l) => l.created_by).filter(Boolean),
      ...state.leads.map((l) => l.assigned_to).filter(Boolean),
      ...state.stages.map((s) => s.created_by).filter(Boolean),
      state.currentUser?.id
    ]);

    const people = state.profiles
      .filter((p) => activeIds.has(p.id))
      .sort((a, b) => String(a.full_name || "").localeCompare(String(b.full_name || "")));

    els.teamList.innerHTML = people.map((profile) => `
      <div class="team-item">
        <strong>${escapeHtml(profile.full_name || "Usuário")}</strong><br>
        E-mail: ${escapeHtml(profile.email || "-")}<br>
        Status: Ativo no CRM compartilhado
      </div>
    `).join("") || '<div class="team-item">Nenhum usuário encontrado.</div>';
  }

  function renderStagesConfig() {
    els.stagesConfigList.innerHTML = state.stages.map((stage, index) => `
      <div class="stage-config-item">
        <div>
          <strong>${escapeHtml(stage.name)}</strong><br>
          <div class="stage-meta-row">
            <span class="status-pill" style="background:${stage.color}22;color:${stage.color};border-color:${stage.color}55">
              ${escapeHtml(stageTypeLabel(stage.stage_type, stage.custom_stage_type))}
            </span>
            <span class="color-chip" title="${escapeHtml(stage.color)}" style="background:${stage.color}"></span>
            <span class="stage-position-label">Posição ${index + 1}</span>
          </div>
        </div>

        <div class="stage-config-actions">
          <button type="button" class="icon-btn" data-stage-action="move-left" data-id="${stage.id}" ${index === 0 ? "disabled" : ""} title="Mover para a esquerda">&larr;</button>
          <button type="button" class="icon-btn" data-stage-action="move-right" data-id="${stage.id}" ${index === state.stages.length - 1 ? "disabled" : ""} title="Mover para a direita">&rarr;</button>
          <button type="button" class="edit-btn" data-stage-action="edit" data-id="${stage.id}">Editar</button>
          <button type="button" class="delete-btn" data-stage-action="delete" data-id="${stage.id}">Excluir</button>
        </div>
      </div>
    `).join("");
  }

  function renderHistoryText() {
    if (!state.history.length) {
      els.historyText.textContent = "Nenhum registro de alteração ainda.";
      return;
    }

    els.historyText.textContent = state.history.map((item) => {
      const when = item.created_at ? new Date(item.created_at).toLocaleString("pt-BR") : "-";
      const who = item.user_name || item.user_email || getProfileNameById(item.user_id) || "Usuário";
      return `[${when}] ${who}\n${item.description || item.action || "Alteração registrada"}\n`;
    }).join("\n");
  }

  function destroyChart(key) {
    if (state.charts[key]) {
      state.charts[key].destroy();
      state.charts[key] = null;
    }
  }

  function makeChartConfig(type, labels, datasets, extraOptions = {}) {
    const textColor = "#edf7f0";
    const gridColor = "rgba(255,255,255,0.08)";
    return {
      type,
      data: { labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        plugins: {
          legend: { labels: { color: textColor } },
          tooltip: { mode: "index", intersect: false }
        },
        scales: {
          x: { ticks: { color: textColor }, grid: { color: gridColor } },
          y: { ticks: { color: textColor }, grid: { color: gridColor }, beginAtZero: true }
        },
        ...extraOptions
      }
    };
  }

  function stageColors() {
    return state.stages.map((s) => s.color || "#1f9d55");
  }

  function renderCharts() {
    if (typeof Chart === "undefined") return;

    ["pipeline", "traffic", "owner", "monthly", "social", "planCount", "planRevenue"].forEach(destroyChart);

    const reportView = $("view-relatorios");
    if (!reportView || !reportView.classList.contains("active-view")) return;

    const filtered = getFilteredLeads();
    const metrics = getDashboardMetrics(filtered);
    const filteredAllMonths = getFilteredLeads({ ignoreMonth: true });
    const qualifiedClosed = getQualifiedClosedLeads(filtered);

    const trafficMap = filtered.reduce((acc, lead) => {
      const key = lead.traffic_type || "Não informado";
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});

    const ownerMap = qualifiedClosed.reduce((acc, lead) => {
      const key = lead.owner || "Sem responsável";
      acc[key] = (acc[key] || 0) + Number(lead.value || 0);
      return acc;
    }, {});

    const monthMap = filteredAllMonths.reduce((acc, lead) => {
      const key = getLeadMonthKey(lead);
      if (!key) return acc;
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});

    const socialMap = filtered.reduce((acc, lead) => {
      const key = lead.social_source || "Não informado";
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});

    const planLabels = metrics.planSummary.map((item) => `${item.plan} - ${brMoney(item.unitValue)}`);

    const create = (key, id, config) => {
      const canvas = $(id);
      if (!canvas) return;
      state.charts[key] = new Chart(canvas, config);
    };

    create("pipeline", "pipelineChart", makeChartConfig("bar", metrics.byStage.map((item) => item.name), [{ label: "Leads", data: metrics.byStage.map((item) => item.count), backgroundColor: stageColors(), borderColor: stageColors() }]));
    create("traffic", "trafficChart", makeChartConfig("doughnut", Object.keys(trafficMap), [{ label: "Origem", data: Object.values(trafficMap) }], { scales: {} }));
    create("owner", "ownerChart", makeChartConfig("bar", Object.keys(ownerMap), [{ label: "Receita fechada", data: Object.values(ownerMap) }], { indexAxis: "y" }));

    const monthLabels = Object.keys(monthMap).sort();
    create("monthly", "monthlyChart", makeChartConfig("line", monthLabels, [{ label: "Leads por mês", data: monthLabels.map((key) => monthMap[key]), tension: 0.3, fill: false }]));

    const socialEntries = Object.entries(socialMap).sort((a, b) => b[1] - a[1]).slice(0, 8);
    create("social", "socialChart", makeChartConfig("bar", socialEntries.map(([key]) => key), [{ label: "Leads", data: socialEntries.map(([, value]) => value) }]));
    create("planCount", "planCountChart", makeChartConfig("bar", planLabels, [{ label: "Fechamentos", data: metrics.planSummary.map((item) => item.count) }]));
    create("planRevenue", "planRevenueChart", makeChartConfig("bar", planLabels, [{ label: "Receita", data: metrics.planSummary.map((item) => item.totalValue) }]));

    setTimeout(() => {
      Object.values(state.charts).forEach((chart) => chart?.resize?.());
    }, 50);
  }


  function normalizeCsvHeader(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function parseCsvLine(line, delimiter) {
    const values = [];
    let current = "";
    let inQuotes = false;

    for (let i = 0; i < line.length; i += 1) {
      const char = line[i];
      const next = line[i + 1];

      if (char === '"') {
        if (inQuotes && next === '"') {
          current += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === delimiter && !inQuotes) {
        values.push(current);
        current = "";
      } else {
        current += char;
      }
    }

    values.push(current);
    return values.map((value) => value.trim());
  }

  function parseCsv(content) {
    const text = String(content || "").replace(/^\uFEFF/, "").trim();
    if (!text) return { headers: [], rows: [] };

    const lines = text.split(/\r?\n/).filter((line) => line.trim());
    const sample = lines.slice(0, 5).join("\n");
    const delimiter = (sample.match(/;/g) || []).length > (sample.match(/,/g) || []).length ? ";" : ",";

    const rawHeaders = parseCsvLine(lines[0], delimiter);
    const headers = rawHeaders.map(normalizeCsvHeader);
    const rows = lines.slice(1).map((line) => {
      const values = parseCsvLine(line, delimiter);
      const row = {};
      headers.forEach((header, index) => {
        row[header] = values[index] || "";
      });
      return row;
    });

    return { headers, rows };
  }

  function csvEscape(value) {
    const text = String(value ?? "");
    if (/[",\n;]/.test(text)) {
      return `"${text.replaceAll('"', '""')}"`;
    }
    return text;
  }

  function stageIdByName(name) {
    const normalized = String(name || "").trim().toLowerCase();
    return state.stages.find((stage) => String(stage.name || "").trim().toLowerCase() === normalized)?.id || null;
  }

  function inferStageId(row) {
    return (
      row.stage_id ||
      stageIdByName(row.pipeline) ||
      stageIdByName(row.etapa) ||
      stageIdByName(row.stage) ||
      state.stages[0]?.id ||
      null
    );
  }

  function parseMoney(value) {
    if (value === null || value === undefined || value === "") return 0;
    const cleaned = String(value)
      .replace(/r\$\s*/gi, "")
      .replace(/\./g, "")
      .replace(",", ".")
      .replace(/[^0-9.-]/g, "");
    return Number(cleaned || 0);
  }

  function downloadTextFile(filename, content, mimeType = "text/plain;charset=utf-8") {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function exportLeadsToCsv() {
    const rows = getFilteredLeads();
    const headers = [
      "nome",
      "contato",
      "responsavel",
      "valor",
      "data_inicio",
      "rede_social",
      "origem",
      "plano",
      "pipeline",
      "observacoes"
    ];

    const csvRows = [
      headers.join(";"),
      ...rows.map((lead) => [
        csvEscape(lead.name || ""),
        csvEscape(lead.contact || ""),
        csvEscape(lead.owner || ""),
        csvEscape(hasLeadValue(lead) ? Number(lead.value || 0).toFixed(2).replace(".", ",") : ""),
        csvEscape(lead.start_date || ""),
        csvEscape(lead.social_source || ""),
        csvEscape(lead.traffic_type || ""),
        csvEscape(getLeadPlan(lead) || ""),
        csvEscape(getStageName(lead.stage_id)),
        csvEscape(lead._meta?.legacyText || "")
      ].join(";"))
    ];

    downloadTextFile("leads.csv", csvRows.join("\n"), "text/csv;charset=utf-8");
  }

  async function importLeadsFromCsv(file) {
    if (!file) return;

    const content = await file.text();
    const { rows } = parseCsv(content);

    if (!rows.length) {
      alert("O CSV está vazio ou inválido.");
      return;
    }

    const payload = rows
      .map((row) => {
        const name = row.nome || row.name || "";
        const contact = row.contato || row.contact || "";
        const owner = row.responsavel || row.owner || getUserDisplayName();
        const startDate = row.data_inicio || row.start_date || "";
        const trafficType = row.origem || row.traffic_type || "Organico";
        const legacyText = String(row.observacoes || row.notes || "").trim();
        const plan = String(row.plano || row.plan || "").trim();

        return {
          stage_id: inferStageId(row),
          assigned_to: state.currentUser?.id || null,
          created_by: state.currentUser?.id || null,
          name: String(name).trim(),
          contact: String(contact).trim(),
          owner: String(owner).trim(),
          value: parseMoney(row.valor || row.value || 0),
          start_date: normalizeDateInput(startDate) || startDate,
          social_source: String(row.rede_social || row.social_source || "").trim(),
          traffic_type: String(trafficType).trim(),
          notes: serializeLeadMeta({
            plan,
            legacyText,
            observations: legacyText ? [{ date: "", text: legacyText }] : []
          })
        };
      })
      .filter((lead) => lead.name && lead.contact && lead.stage_id);

    if (!payload.length) {
      alert("Nenhuma linha válida encontrada. Use colunas como nome, contato, valor, data_inicio e pipeline.");
      return;
    }

    const { data, error } = await state.supabase
      .from("leads")
      .insert(payload)
      .select("id,name");

    if (error) {
      alert(error.message);
      return;
    }

    await logChange(
      "import_csv",
      "lead",
      null,
      `${payload.length} lead(s) foram importados via CSV por ${getUserDisplayName()}.`,
      { imported_count: payload.length, imported_ids: (data || []).map((item) => item.id) }
    );

    alert(`${payload.length} lead(s) importado(s) com sucesso.`);
    els.csvFileInput.value = "";
    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  function renderAll() {
    syncSelectedLeadIds();
    populateFilters();
    renderStats();
    renderPipeline();
    renderLeadTable();
    renderTeam();
    renderStagesConfig();
    renderHistoryText();
    if (state.activeView === "relatorios" && typeof window.Chart !== "undefined") {
      renderCharts();
    }
    bindGeneralActionEvents();
  }

  function openLeadModal(lead = null) {
    els.leadForm.reset();
    els.leadId.value = lead?.id || "";
    els.modalTitle.textContent = lead ? "Editar Lead" : "Novo Lead";

    els.name.value = lead?.name || "";
    els.contact.value = lead?.contact || "";
    els.owner.value = lead?.owner || getUserDisplayName();
    els.value.value = lead?.value || "";
    els.startDate.value = lead?.start_date || "";
    els.stage.value = lead?.stage_id || state.stages[0]?.id || "";
    if (!state.stages.length) {
      alert("Nenhuma etapa encontrada no CRM. Crie uma etapa primeiro.");
      closeLeadModal();
      return;
    }
    els.socialSource.value = lead?.social_source || "";
    els.trafficType.value = lead?.traffic_type || "Organico";
    state.modalPlans = getLeadPlans(lead).map((item) => ({ ...item }));
    state.modalObservations = getLeadObservations(lead);
    renderPlanItems();
    renderObservationItems();

    els.modalOverlay.classList.remove("hidden");
  }

  function closeLeadModal() {
    els.modalOverlay.classList.add("hidden");
  }

  function openStageModal(stage = null) {
    els.stageForm.reset();
    els.stageId.value = stage?.id || "";
    els.stageModalTitle.textContent = stage ? "Editar pipeline" : "Adicionar pipeline";
    els.saveStageBtn.textContent = stage ? "Salvar alterações" : "Adicionar";
    els.stageName.value = stage?.name || "";
    els.stageColor.value = stage?.color || "#1f9d55";
    refreshStageTypeOptions(stage?.custom_stage_type ? `custom:${stage.custom_stage_type}` : (stage?.stage_type || "andamento"), stage?.custom_stage_type || "");
    updateStageColorPreview(els.stageColor.value);
    els.stageModalOverlay.classList.remove("hidden");
  }

  function closeStageModal() {
    els.stageModalOverlay.classList.add("hidden");
  }

  async function openHistoryModal() {
    els.historyText.textContent = state.historyLoaded ? els.historyText.textContent : "Carregando histórico...";
    renderHistoryText();
    els.historyModalOverlay.classList.remove("hidden");
    await loadHistory();
  }

  function closeHistoryModal() {
    els.historyModalOverlay.classList.add("hidden");
  }

  async function logChange(action, entityType, entityId, description, payload = null) {
    if (!state.currentUser) return;

    const row = {
      action,
      entity_type: entityType,
      entity_id: entityId ? String(entityId) : null,
      description,
      payload: payload || {},
      user_id: state.currentUser.id,
      user_name: getUserDisplayName(),
      user_email: state.currentUser.email
    };

    const { error } = await state.supabase.from("change_history").insert([row]);
    if (error) console.error("Erro ao gravar histórico:", error);
    state.historyLoaded = false;
  }

  async function submitLead(event) {
    event.preventDefault();

    const existingLead = state.leads.find((lead) => lead.id === els.leadId.value) || null;
    const existingMeta = getLeadMeta(existingLead?.notes || "");
    const rawValue = String(els.value.value || "").trim();
    const parsedLeadValue = parseMoney(rawValue);
    const leadValue = Number.isFinite(parsedLeadValue) ? Number(parsedLeadValue) : 0;
    const draftPlans = cleanPlanList(state.modalPlans.map((item) => ({ name: item.name, value: item.value })));
    const draftObservations = cleanObservationList(state.modalObservations);

    const invalidPlan = state.modalPlans.find((item) => String(item?.name || "").trim() && String(item?.value || "").trim() === "");
    if (invalidPlan) return alert("Ao adicionar um plano, informe também o valor.");

    const payload = {
      assigned_to: state.currentUser.id,
      created_by: existingLead?.created_by || state.currentUser.id,
      stage_id: els.stage.value,
      name: els.name.value.trim(),
      contact: els.contact.value.trim(),
      owner: els.owner.value.trim(),
      value: Number.isFinite(leadValue) ? leadValue : 0,
      start_date: normalizeDateInput(els.startDate.value) || els.startDate.value,
      social_source: els.socialSource.value.trim(),
      traffic_type: els.trafficType.value,
      notes: serializeLeadMeta({
        ...existingMeta,
        plans: draftPlans,
        plan: draftPlans[0]?.name || "",
        legacyText: existingMeta.legacyText,
        observations: draftObservations
      })
    };

    if (!payload.stage_id) return alert("Selecione uma etapa.");
    if (!payload.name || !payload.contact || !payload.owner || !payload.start_date) {
      return alert("Preencha os campos obrigatórios.");
    }

    let error;

    if (els.leadId.value) {
      const oldLead = state.leads.find((x) => x.id === els.leadId.value);

      ({ error } = await state.supabase
        .from("leads")
        .update(payload)
        .eq("id", els.leadId.value));

      if (error) return alert(`Erro no Supabase: ${error.message}`);

      await logChange(
        "update",
        "lead",
        els.leadId.value,
        `Lead "${payload.name}" foi atualizado por ${getUserDisplayName()}.`,
        { before: oldLead || null, after: payload }
      );
    } else {
      const { data, error: insertError } = await state.supabase
        .from("leads")
        .insert([payload])
        .select()
        .single();

      error = insertError;
      if (error) return alert(`Erro no Supabase: ${error.message}`);

      await logChange(
        "insert",
        "lead",
        data?.id,
        `Lead "${payload.name}" foi criado por ${getUserDisplayName()}.`,
        payload
      );
    }

    closeLeadModal();
    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  async function submitStage(event) {
    event.preventDefault();

    const selectedType = String(els.stageType.value || "andamento").trim();

    const customStageType = selectedType.startsWith("custom:") ? selectedType.replace(/^custom:/, "") : String(els.customStageType?.value || "").trim();
    const payload = {
      name: els.stageName.value.trim(),
      color: els.stageColor.value,
      stage_type: (selectedType === "personalizado" || selectedType.startsWith("custom:")) ? "personalizado" : (["andamento", "fechado", "cancelado", "espera"].includes(selectedType) ? selectedType : "andamento"),
      custom_stage_type: (selectedType === "personalizado" || selectedType.startsWith("custom:")) ? customStageType : null,
      position: els.stageId.value ? (state.stages.find((s) => s.id === els.stageId.value)?.position ?? state.stages.length) : state.stages.length,
      created_by: state.currentUser.id
    };

    if (!payload.name) return alert("Informe o nome da etapa.");
    if (payload.stage_type === "personalizado" && !payload.custom_stage_type) return alert("Informe o tipo personalizado.");

    if (payload.stage_type === "personalizado") {
      const savedType = await saveCustomStageType(payload.custom_stage_type);
      if (!savedType) return alert("Não foi possível salvar o novo tipo personalizado.");
    }

    if (els.stageId.value) {
      const before = state.stages.find((s) => s.id === els.stageId.value);
      const { error } = await state.supabase.from("stages").update({ name: payload.name, color: payload.color, stage_type: payload.stage_type, custom_stage_type: payload.custom_stage_type, position: payload.position }).eq("id", els.stageId.value);
      if (error) return alert(`Erro no Supabase: ${error.message}`);
      await logChange("update", "stage", els.stageId.value, `Etapa "${before?.name || payload.name}" foi atualizada por ${getUserDisplayName()}.`, { before, after: payload });
    } else {
      const { data, error } = await state.supabase.from("stages").insert([payload]).select().single();
      if (error) return alert(`Erro no Supabase: ${error.message}`);
      await logChange("insert", "stage", data?.id, `Etapa "${payload.name}" foi criada por ${getUserDisplayName()}.`, payload);
    }

    closeStageModal();
    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  async function moveLeadToStage(leadId, stageId) {
    const lead = state.leads.find((x) => x.id === leadId);
    const stage = state.stages.find((x) => x.id === stageId);
    const fromStage = state.stages.find((x) => x.id === lead?.stage_id);

    if (!lead || !stage || lead.stage_id === stageId) return;

    const { error } = await state.supabase
      .from("leads")
      .update({ stage_id: stageId })
      .eq("id", lead.id);

    if (error) return alert(`Erro no Supabase: ${error.message}`);

    await logChange(
      "move_stage",
      "lead",
      lead.id,
      `Lead "${lead.name}" foi movido de "${fromStage?.name || "-"}" para "${stage.name}" por ${getUserDisplayName()}.`,
      {
        lead_id: lead.id,
        from_stage_id: fromStage?.id || null,
        from_stage_name: fromStage?.name || null,
        to_stage_id: stage.id,
        to_stage_name: stage.name
      }
    );

    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  async function deleteLead(id) {
    const lead = state.leads.find((x) => x.id === id);
    if (!lead) return;
    if (!confirm(`Excluir o lead "${lead.name}"?`)) return;

    const { error } = await state.supabase
      .from("leads")
      .delete()
      .eq("id", id);

    if (error) return alert(`Erro no Supabase: ${error.message}`);

    await logChange(
      "delete",
      "lead",
      id,
      `Lead "${lead.name}" foi excluído por ${getUserDisplayName()}.`,
      lead
    );

    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  async function editStage(id) {
    const stage = state.stages.find((x) => x.id === id);
    if (!stage) return;
    openStageModal(stage);
  }

  async function deleteStage(id) {
    const stage = state.stages.find((x) => x.id === id);
    const fallback = state.stages.find((x) => x.id !== id);

    if (!stage) return;
    if (!fallback) return alert("Você precisa manter ao menos uma etapa.");
    if (!confirm(`Excluir a etapa "${stage.name}"?`)) return;

    const affectedLeads = state.leads.filter((lead) => lead.stage_id === id);

    if (affectedLeads.length) {
      const { error: leadError } = await state.supabase
        .from("leads")
        .update({ stage_id: fallback.id })
        .eq("stage_id", id);

      if (leadError) return alert(leadError.message);

      await logChange(
        "bulk_move_stage",
        "lead",
        null,
        `${affectedLeads.length} lead(s) foram movidos automaticamente de "${stage.name}" para "${fallback.name}" porque a etapa foi excluída por ${getUserDisplayName()}.`,
        {
          from_stage_id: stage.id,
          from_stage_name: stage.name,
          to_stage_id: fallback.id,
          to_stage_name: fallback.name,
          affected_count: affectedLeads.length
        }
      );
    }

    const { error } = await state.supabase
      .from("stages")
      .delete()
      .eq("id", id);

    if (error) return alert(`Erro no Supabase: ${error.message}`);

    await logChange(
      "delete",
      "stage",
      stage.id,
      `Etapa "${stage.name}" foi excluída por ${getUserDisplayName()}.`,
      stage
    );

    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  async function sendResetPasswordEmail() {
    const email = $("loginEmail").value.trim();
    if (!email) return setMessage(els.authMessage, "Digite seu e-mail para recuperar a senha.", true);

    const redirectTo = `${window.location.origin}${window.location.pathname}#type=recovery`;

    const { error } = await state.supabase.auth.resetPasswordForEmail(email, { redirectTo });
    if (error) return setMessage(els.authMessage, error.message, true);

    setMessage(els.authMessage, "Enviamos o link de recuperação. Verifique seu e-mail.");
  }

  async function updateRecoveredPassword() {
    const newPassword = $("newPassword").value.trim();
    if (!newPassword) return setMessage(els.authMessage, "Digite a nova senha.", true);

    const { error } = await state.supabase.auth.updateUser({ password: newPassword });
    if (error) return setMessage(els.authMessage, error.message, true);

    els.resetPasswordBox.classList.add("hidden");
    setMessage(els.authMessage, "Senha atualizada com sucesso. Faça login.");
    window.location.hash = "";
    await state.supabase.auth.signOut();
    showScreen("authScreen");
  }

  function bindView(name) {
    state.activeView = name;

    document.querySelectorAll(".menu-item").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.view === name);
    });

    document.querySelectorAll(".view").forEach((view) => {
      view.classList.toggle("active-view", view.id === `view-${name}`);
    });

    const titles = {
      funil: ["Pipeline de Vendas", "Gerencie os leads por etapa compartilhada."],
      leads: ["Lista de Leads", "Visualize todos os leads cadastrados."],
      relatorios: ["Relatórios", "Acompanhe os resultados do CRM compartilhado."],
      equipe: ["Equipe", "Usuários que já interagiram no CRM."],
      configuracoes: ["Configurações", "Gerencie as etapas do pipeline compartilhado."]
    };

    els.pageTitle.textContent = titles[name][0];
    els.pageSubtitle.textContent = titles[name][1];
    els.sidebar.classList.remove("open");

    if (name === "relatorios") {
      setTimeout(async () => {
        try {
          await ensureChartLibrary();
          renderCharts();
        } catch (error) {
          console.error(error);
        }
      }, 80);
    }

    if (name === "equipe") {
      void loadProfilesIfNeeded();
    }
  }

  function bindPipelineEvents() {
    document.querySelectorAll(".card").forEach((card) => {
      card.ondragstart = () => card.classList.add("dragging");
      card.ondragend = () => card.classList.remove("dragging");
    });

    document.querySelectorAll(".column").forEach((column) => {
      column.ondragover = (e) => {
        e.preventDefault();
        column.classList.add("drag-over");
      };

      column.ondragleave = () => {
        column.classList.remove("drag-over");
      };

      column.ondrop = async (e) => {
        e.preventDefault();
        column.classList.remove("drag-over");
        const dragged = document.querySelector(".card.dragging");
        const leadId = dragged?.dataset?.leadId;
        const stageId = column.dataset.stageId;
        if (leadId && stageId) await moveLeadToStage(leadId, stageId);
      };
    });

    document.querySelectorAll('#view-funil [data-action="edit-lead"]').forEach((btn) => {
      btn.onclick = () => {
        const lead = state.leads.find((x) => x.id === btn.dataset.id);
        if (lead) openLeadModal(lead);
      };
    });

    document.querySelectorAll('#view-funil [data-action="delete-lead"]').forEach((btn) => {
      btn.onclick = () => deleteLead(btn.dataset.id);
    });
  }

  function bindGeneralActionEvents() {
    document.querySelectorAll('#view-leads [data-action="edit-lead"]').forEach((btn) => {
      btn.onclick = () => {
        const lead = state.leads.find((x) => x.id === btn.dataset.id);
        if (lead) openLeadModal(lead);
      };
    });

    document.querySelectorAll('#view-leads [data-action="delete-lead"]').forEach((btn) => {
      btn.onclick = () => deleteLead(btn.dataset.id);
    });

    document.querySelectorAll('[data-stage-action="edit"]').forEach((btn) => {
      btn.onclick = () => editStage(btn.dataset.id);
    });

    document.querySelectorAll('[data-stage-action="delete"]').forEach((btn) => {
      btn.onclick = () => deleteStage(btn.dataset.id);
    });

    document.querySelectorAll('[data-stage-action="move-left"]').forEach((btn) => {
      btn.onclick = () => moveStagePosition(btn.dataset.id, "left");
    });

    document.querySelectorAll('[data-stage-action="move-right"]').forEach((btn) => {
      btn.onclick = () => moveStagePosition(btn.dataset.id, "right");
    });
  }

  function bindEvents() {
    document.querySelectorAll(".tab-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".tab-btn").forEach((b) => b.classList.toggle("active", b === btn));
        document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.remove("active"));
        $(`${btn.dataset.tab}Form`).classList.add("active");
      });
    });

    els.loginForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      setMessage(els.authMessage, "");

      const { data, error } = await state.supabase.auth.signInWithPassword({
        email: $("loginEmail").value.trim(),
        password: $("loginPassword").value.trim()
      });

      if (error) return setMessage(els.authMessage, "E-mail ou senha inválidos.", true);

      state.currentUser = data?.user || null;
      if (!state.currentUser) {
        return setMessage(els.authMessage, "Não foi possível iniciar a sessão. Tente novamente.", true);
      }

      await ensureProfile();
      await enterApp();
    });

    els.registerForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      setMessage(els.authMessage, "");

      const email = $("registerEmail").value.trim();
      const password = $("registerPassword").value.trim();
      const full_name = $("registerName").value.trim();

      const { data, error } = await state.supabase.auth.signUp({
        email,
        password,
        options: {
          data: { full_name }
        }
      });

      if (error) return setMessage(els.authMessage, error.message, true);

      if (data.user) {
        state.currentUser = data.user;
        await ensureProfile();
      }

      setMessage(els.authMessage, "Conta criada com sucesso. Agora faça login.");
      document.querySelector('[data-tab="login"]').click();
      $("loginEmail").value = email;
      $("registerForm").reset();
    });

    els.forgotPasswordBtn.addEventListener("click", sendResetPasswordEmail);
    els.updatePasswordBtn.addEventListener("click", updateRecoveredPassword);

    els.logoutBtn.addEventListener("click", async () => {
      await state.supabase.auth.signOut();
      state.currentUser = null;
      state.profile = null;
      state.stages = [];
      state.customStageTypes = [];
      state.leads = [];
      state.profiles = [];
      state.history = [];
      showScreen("authScreen");
    });

    els.mobileMenuBtn.addEventListener("click", () => {
      els.sidebar.classList.toggle("open");
    });

    document.querySelectorAll(".menu-item").forEach((btn) => {
      btn.addEventListener("click", () => bindView(btn.dataset.view));
    });

    els.searchInput.addEventListener("input", () => {
      renderAll();
    });

    els.ownerFilterBtn?.addEventListener("click", (e) => {
      e.stopPropagation();
      const isOpen = els.ownerFilterDropdown?.classList.contains("open");
      closeFilterDropdowns(els.ownerFilterDropdown);
      setFilterDropdownOpen(els.ownerFilterDropdown, els.ownerFilterBtn, els.ownerFilterMenu, !isOpen);
    });

    els.monthFilterBtn?.addEventListener("click", (e) => {
      e.stopPropagation();
      const isOpen = els.monthFilterDropdown?.classList.contains("open");
      closeFilterDropdowns(els.monthFilterDropdown);
      setFilterDropdownOpen(els.monthFilterDropdown, els.monthFilterBtn, els.monthFilterMenu, !isOpen);
    });

    document.addEventListener("click", (e) => {
      if (els.ownerFilterDropdown?.contains(e.target) || els.monthFilterDropdown?.contains(e.target)) return;
      closeFilterDropdowns();
    });

    els.deleteSelectedBtn?.addEventListener("click", deleteSelectedLeads);

    els.selectAllLeads?.addEventListener("change", () => {
      const filtered = getFilteredLeads();
      const visibleIds = filtered.map((lead) => lead.id);

      if (els.selectAllLeads.checked) {
        visibleIds.forEach((id) => state.selectedLeadIds.add(id));
      } else {
        visibleIds.forEach((id) => state.selectedLeadIds.delete(id));
      }

      renderLeadTable();
    });

    els.importCsvBtn.addEventListener("click", () => els.csvFileInput.click());
    els.exportCsvBtn.addEventListener("click", exportLeadsToCsv);
    els.csvFileInput.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      await importLeadsFromCsv(file);
    });

    els.addLeadBtn.addEventListener("click", () => openLeadModal());
    els.addPlanBtn?.addEventListener("click", addPlanFromDraft);
    els.addObservationBtn?.addEventListener("click", addObservationFromDraft);
    els.addStageBtn.addEventListener("click", openStageModal);
    els.historyBtn.addEventListener("click", openHistoryModal);

    els.closeModalBtn.addEventListener("click", closeLeadModal);
    els.cancelBtn.addEventListener("click", closeLeadModal);
    els.leadForm.addEventListener("submit", submitLead);

    els.closeStageModalBtn.addEventListener("click", closeStageModal);
    els.cancelStageBtn.addEventListener("click", closeStageModal);
    els.stageForm.addEventListener("submit", submitStage);
    els.stageColor.addEventListener("input", (e) => updateStageColorPreview(e.target.value));
    els.stageType.addEventListener("change", toggleCustomStageTypeField);
    els.removeCustomTypeBtn.addEventListener("click", removeCurrentCustomStageType);
    els.savedStageTypes?.addEventListener("click", async (event) => {
      const selectBtn = event.target.closest("[data-stage-type]");
      if (selectBtn) {
        const type = String(selectBtn.dataset.stageType || "").trim();
        if (!type) return;
        els.stageType.value = `custom:${type}`;
        toggleCustomStageTypeField();
        return;
      }

      const deleteBtn = event.target.closest("[data-stage-type-delete]");
      if (!deleteBtn) return;

      const type = String(deleteBtn.dataset.stageTypeDelete || "").trim();
      if (!type) return;

      els.stageType.value = `custom:${type}`;
      await removeCurrentCustomStageType();
    });
    setupPlanListEvents();
    setupObservationListEvents();
    attachPipelineScrollEvents();

    document.addEventListener("click", (event) => {
      const leadBtn = event.target.closest("[data-action]");
      if (leadBtn) {
        if (leadBtn.dataset.action === "edit-lead") {
          const lead = state.leads.find((x) => x.id === leadBtn.dataset.id);
          if (lead) openLeadModal(lead);
        }
        if (leadBtn.dataset.action === "delete-lead") deleteLead(leadBtn.dataset.id);
      }

      const stageBtn = event.target.closest("[data-stage-action]");
      if (!stageBtn) return;
      const stageId = stageBtn.dataset.id;
      if (stageBtn.dataset.stageAction === "edit") editStage(stageId);
      if (stageBtn.dataset.stageAction === "delete") deleteStage(stageId);
      if (stageBtn.dataset.stageAction === "move-left") moveStagePosition(stageId, "left");
      if (stageBtn.dataset.stageAction === "move-right") moveStagePosition(stageId, "right");
    });

    els.closeHistoryModalBtn.addEventListener("click", closeHistoryModal);
    els.refreshHistoryBtn.addEventListener("click", async () => {
      els.historyText.textContent = "Carregando histórico...";
      await loadHistory(true);
    });

    els.modalOverlay.addEventListener("click", (e) => {
      if (e.target === els.modalOverlay) closeLeadModal();
    });

    els.stageModalOverlay.addEventListener("click", (e) => {
      if (e.target === els.stageModalOverlay) closeStageModal();
    });

    els.historyModalOverlay.addEventListener("click", (e) => {
      if (e.target === els.historyModalOverlay) closeHistoryModal();
    });

    state.supabase.auth.onAuthStateChange(async (event, session) => {
      if (event === "PASSWORD_RECOVERY") {
        els.resetPasswordBox.classList.remove("hidden");
        showScreen("authScreen");
        setMessage(els.authMessage, "Digite sua nova senha e salve.");
        return;
      }

      if (event === "SIGNED_OUT") {
        state.currentUser = null;
        state.profile = null;
        state.stages = [];
        state.customStageTypes = [];
        state.leads = [];
        state.profiles = [];
        state.history = [];
        state.activeView = "funil";
        state.historyLoaded = false;
        state.profilesLoaded = false;
        showScreen("authScreen");
        return;
      }

      if (event === "SIGNED_IN" && session?.user) {
        state.currentUser = session.user;
        await ensureProfile();
        if (els.appScreen.classList.contains("hidden")) {
          await enterApp();
        } else {
          await loadAppData({ includeProfiles: state.profilesLoaded });
        }
      }
    });
  }

  async function init() {
    try {
      createClient();
      bindEvents();
      await bootstrap();
    } catch (error) {
      console.error(error);
      els.bootScreen.textContent = "Erro ao iniciar o sistema. Verifique o config.js e o Supabase.";
    }
  }

  init();
})();
