(() => {
  const $ = (id) => document.getElementById(id);

  const els = {
    bootScreen: $("bootScreen"),
    authScreen: $("authScreen"),
    appScreen: $("appScreen"),
    authMessage: $("authMessage"),
    registerTabBtn: $("registerTabBtn"),
    registrationNotice: $("registrationNotice"),
    resetPasswordBox: $("resetPasswordBox"),
    loginForm: $("loginForm"),
    registerForm: $("registerForm"),
    forgotPasswordBtn: $("forgotPasswordBtn"),
    updatePasswordBtn: $("updatePasswordBtn"),
    logoutBtn: $("logoutBtn"),
    mobileMenuBtn: $("mobileMenuBtn"),
    mobileTopbar: document.querySelector(".mobile-topbar"),
    sidebar: $("sidebar"),
    main: document.querySelector(".main"),
    topbar: document.querySelector(".topbar"),

    orgNameLabel: $("orgNameLabel"),
    mobileOrgName: $("mobileOrgName"),
    userWelcome: $("userWelcome"),
    userAccessLevel: $("userAccessLevel"),
    pageTitle: $("pageTitle"),
    pageSubtitle: $("pageSubtitle"),

    searchInput: $("searchInput"),
    mobileFiltersBtn: $("mobileFiltersBtn"),
    mobileFiltersPanel: $("mobileFiltersPanel"),
    mobileOwnerFilter: $("mobileOwnerFilter"),
    mobileMonthFilter: $("mobileMonthFilter"),
    mobileLeadSourceFilter: $("mobileLeadSourceFilter"),
    mobileClearFiltersBtn: $("mobileClearFiltersBtn"),
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
    leadSourceFilterDropdown: $("leadSourceFilterDropdown"),
    leadSourceFilterBtn: $("leadSourceFilterBtn"),
    leadSourceFilterMenu: $("leadSourceFilterMenu"),
    leadSourceFilterLabel: $("leadSourceFilterLabel"),
    leadSourceFilter: $("leadSourceFilter"),
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
    pipelineScrollTopTrack: $("pipelineScrollTopTrack"),
    pipelineScrollTopThumb: $("pipelineScrollTopThumb"),
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
    metricsSection: $("metricsSection"),
    funilStickyHead: $("funilStickyHead"),

    reportClosedValue: $("reportClosedValue"),
    reportTotalLeads: $("reportTotalLeads"),
    reportClosedDeals: $("reportClosedDeals"),
    reportConversionRate: $("reportConversionRate"),
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
    planSuggestions: $("planSuggestions"),
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
    ownerGroup: $("ownerGroup"),
    planGroup: $("planGroup"),
    customStageTypeGroup: $("customStageTypeGroup"),
    customStageType: $("customStageType"),
    removeCustomTypeBtn: $("removeCustomTypeBtn"),
    savedStageTypesGroup: $("savedStageTypesGroup"),
    savedStageTypes: $("savedStageTypes"),
    savedStageTypeActions: $("savedStageTypeActions"),
    removeSelectedStageTypeBtn: $("removeSelectedStageTypeBtn"),
    addLeadSourceBtn: $("addLeadSourceBtn"),
    leadSourceName: $("leadSourceName"),
    leadSourcesConfigList: $("leadSourcesConfigList"),
    teamAccessLegend: $("teamAccessLegend"),
    accessRequestsList: $("accessRequestsList"),
    adminRequestsList: $("adminRequestsList"),

    historyModalOverlay: $("historyModalOverlay"),
    closeHistoryModalBtn: $("closeHistoryModalBtn"),
    refreshHistoryBtn: $("refreshHistoryBtn"),
    historyText: $("historyText"),

    permissionModalOverlay: $("permissionModalOverlay"),
    closePermissionModalBtn: $("closePermissionModalBtn"),
    cancelPermissionRequestBtn: $("cancelPermissionRequestBtn"),
    submitPermissionRequestBtn: $("submitPermissionRequestBtn"),
    permissionRequestReason: $("permissionRequestReason"),
    permissionModalTitle: $("permissionModalTitle"),
    permissionModalText: $("permissionModalText")
  };

  const state = {
    supabase: null,
    currentUser: null,
    profile: null,
    stages: [],
    customStageTypes: [],
    hiddenPresetStageTypes: [],
    leads: [],
    profiles: [],
    accessRequests: [],
    adminRequests: [],
    leadSources: [],
    history: [],
    activeView: "funil",
    historyLoaded: false,
    profilesLoaded: false,
    permissionRequestContext: null,
    security: {
      allowSelfRegistration: false,
      allowedSignupEmailDomains: []
    },
    chartLoader: null,
    selectedLeadIds: new Set(),
    modalPlans: [],
    modalObservations: [],
    charts: {
      pipeline: null,
      traffic: null,
      owner: null,
      yearlyDaily: null,
      monthly: null,
      social: null,
      planCount: null,
      planRevenue: null
    },
    pipelineScrollObserver: null,
    pipelineScrollbarDrag: null
  };

  const PRESET_STAGE_TYPES = [
    { value: "andamento", label: "Andamento" },
    { value: "fechado", label: "Fechado" },
    { value: "cancelado", label: "Cancelado" },
    { value: "espera", label: "Espera" }
  ];

  const DEFAULT_STAGE_COLOR = "#1F9D55";
  const CHART_JS_URL = "https://cdn.jsdelivr.net/npm/chart.js/dist/chart.umd.min.js";
  const ALLOWED_EXTERNAL_SCRIPT_URLS = new Set([CHART_JS_URL]);
  const USER_ROLE = {
    ADMIN: "admin",
    USER: "user"
  };
  const ACCESS_STATUS = {
    PENDING: "pending",
    APPROVED: "approved",
    REJECTED: "rejected"
  };
  const DEFAULT_LEAD_SOURCES = ["Organico", "Pago"];

  const STORAGE_CACHE_KEYS = [
    "crmPax.customStageTypes",
    "crmPax.hiddenPresetStageTypes"
  ];

  const STORAGE_CLEANUP_KEY = "crmPax.storageCleanupAt";
  const STORAGE_CLEANUP_INTERVAL_MS = 3 * 24 * 60 * 60 * 1000;

  function createClient() {
    const cfg = window.APP_CONFIG || {};
    if (!cfg.supabaseUrl || !cfg.supabaseAnonKey) {
      throw new Error("Configuração do Supabase não encontrada em config.js");
    }
    state.supabase = window.supabase.createClient(cfg.supabaseUrl, cfg.supabaseAnonKey);
  }

  function getSecurityConfig() {
    const cfg = window.APP_CONFIG || {};
    const allowedSignupEmailDomains = [...new Set(
      (Array.isArray(cfg.allowedSignupEmailDomains) ? cfg.allowedSignupEmailDomains : [])
        .map((item) => String(item || "").trim().toLowerCase().replace(/^@+/, ""))
        .filter((item) => item && item.includes(".") && !item.includes(" "))
    )];

    return {
      allowSelfRegistration: cfg.allowSelfRegistration === true,
      allowedSignupEmailDomains
    };
  }

  function getSignupRestrictionMessage() {
    if (!state.security.allowSelfRegistration) {
      return "Novas solicitacoes publicas estao desativadas. Solicite a liberacao diretamente a um administrador.";
    }
    if (!state.security.allowedSignupEmailDomains.length) {
      return "O acesso depende de aprovacao do administrador. Sua solicitacao fica pendente ate a liberacao.";
    }
    return `Solicitacoes aceitas apenas para e-mails de: ${state.security.allowedSignupEmailDomains.join(", ")}.`;
  }

  function applySecurityConfigToUi() {
    const registrationEnabled = state.security.allowSelfRegistration;

    els.registerTabBtn?.classList.toggle("hidden", !registrationEnabled);
    els.registerForm?.classList.toggle("hidden", !registrationEnabled);
    els.registerForm?.classList.remove("active");
    els.loginForm?.classList.add("active");
    document.querySelector('[data-tab="login"]')?.classList.add("active");
    els.registerTabBtn?.classList.remove("active");

    const notice = getSignupRestrictionMessage();
    if (els.registrationNotice) {
      els.registrationNotice.textContent = notice;
      els.registrationNotice.classList.toggle("hidden", !notice);
    }
  }

  function setMessage(el, text, isError = false) {
    el.textContent = text || "";
    el.style.color = isError ? "#fecaca" : "#cdecd6";
  }

  function getAuthErrorMessage(error, fallback = "Nao foi possivel concluir a autenticacao.") {
    const code = String(error?.code || "").toLowerCase();
    const message = String(error?.message || "").toLowerCase();

    if (code === "email_not_confirmed" || message.includes("email not confirmed")) {
      return "Seu e-mail ainda nao foi confirmado. Verifique a caixa de entrada antes de entrar.";
    }

    if (code === "email_provider_disabled" || message.includes("email provider is disabled")) {
      return "O login e cadastro por e-mail estao desabilitados no Supabase. Ative o provedor Email nas configuracoes de Authentication.";
    }

    if (code === "provider_disabled" || message.includes("provider is disabled")) {
      return "Este metodo de autenticacao esta desabilitado no Supabase. Revise Authentication > Providers.";
    }

    if (code === "invalid_credentials" || message.includes("invalid login credentials")) {
      return "E-mail ou senha invalidos.";
    }

    if (code === "email_exists" || message.includes("user already registered")) {
      return "Este e-mail ja esta cadastrado. Tente entrar ou recuperar a senha.";
    }

    if (
      code === "over_email_send_rate_limit" ||
      message.includes("email rate limit exceeded") ||
      message.includes("rate limit exceeded")
    ) {
      return "O Supabase atingiu o limite temporario de envio de e-mails. Aguarde alguns minutos e tente novamente, ou configure um SMTP proprio no projeto.";
    }

    if (code === "email_address_not_authorized" || message.includes("email address not authorized")) {
      return "O Supabase nao vai enviar e-mails para esse endereco usando o SMTP padrao. Configure um SMTP proprio ou teste com um e-mail da equipe do projeto.";
    }

    return error?.message || fallback;
  }

  function getAuthRedirectUrl(hash = "") {
    const protocol = String(window.location.protocol || "").toLowerCase();
    if (protocol !== "http:" && protocol !== "https:") return null;

    const baseUrl = `${window.location.origin}${window.location.pathname}`;
    return hash ? `${baseUrl}${hash}` : baseUrl;
  }

  function getUserRole() {
    return String(state.profile?.role || USER_ROLE.USER).trim().toLowerCase();
  }

  function getAccessStatus() {
    return String(state.profile?.access_status || ACCESS_STATUS.PENDING).trim().toLowerCase();
  }

  function isAdmin() {
    return getUserRole() === USER_ROLE.ADMIN;
  }

  function isApprovedUser() {
    return getAccessStatus() === ACCESS_STATUS.APPROVED;
  }

  function canDeleteLeads() {
    return isAdmin();
  }

  function canAssignLeadOwner() {
    return isAdmin();
  }

  function canManageAdminAreas() {
    return isAdmin();
  }

  function canViewHistory() {
    return isAdmin();
  }

  function canManageStages() {
    return isAdmin();
  }

  function canManageLeadSources() {
    return isAdmin();
  }

  function getRoleLabel(role = getUserRole()) {
    return role === USER_ROLE.ADMIN ? "Administrador" : "Usuario comum";
  }

  function getAccessStatusLabel(status = getAccessStatus()) {
    if (status === ACCESS_STATUS.APPROVED) return "Aprovado";
    if (status === ACCESS_STATUS.REJECTED) return "Recusado";
    return "Pendente";
  }

  function getAccessSummaryLabel() {
    const roleLabel = getRoleLabel();
    const statusLabel = getAccessStatusLabel();
    if (isApprovedUser()) return `${roleLabel} liberado`;
    return `${roleLabel} - ${statusLabel.toLowerCase()}`;
  }

  function applyRoleBasedUi() {
    const adminVisible = canManageAdminAreas();
    document.querySelectorAll("[data-admin-only='true']").forEach((element) => {
      element.classList.toggle("hidden", !adminVisible);
    });

    if (els.userAccessLevel) {
      els.userAccessLevel.textContent = getAccessSummaryLabel();
    }

    if (els.teamAccessLegend) {
      els.teamAccessLegend.classList.toggle("hidden", !adminVisible);
    }
  }

  async function forceSignOutWithMessage(message) {
    resetAppState();
    await state.supabase.auth.signOut();
    showScreen("authScreen");
    setMessage(els.authMessage, message, true);
  }

  function showScreen(id) {
    closeAllModals();
    [els.authScreen, els.appScreen].forEach((screen) => screen.classList.add("hidden"));
    $(id).classList.remove("hidden");
    els.bootScreen.classList.add("hidden");
  }

  function resetAppState() {
    state.currentUser = null;
    state.profile = null;
    state.stages = [];
    state.customStageTypes = [];
    state.hiddenPresetStageTypes = [];
    state.leads = [];
    state.profiles = [];
    state.accessRequests = [];
    state.adminRequests = [];
    state.leadSources = [];
    state.history = [];
    state.activeView = "funil";
    state.historyLoaded = false;
    state.profilesLoaded = false;
    state.permissionRequestContext = null;
    els.resetPasswordBox?.classList.add("hidden");
  }

  function closeAllModals() {
    [els.modalOverlay, els.stageModalOverlay, els.historyModalOverlay, els.permissionModalOverlay].forEach((overlay) => {
      overlay?.classList.add("hidden");
    });
    document.body.classList.remove("modal-open");
  }

  function openModalOverlay(overlay, focusSelector = null) {
    if (!overlay) return;
    overlay.classList.remove("hidden");
    overlay.scrollTop = 0;
    document.body.classList.add("modal-open");

    const modal = overlay.querySelector(".crm-modal");
    if (modal) {
      modal.scrollTop = 0;
      requestAnimationFrame(() => {
        modal.scrollIntoView({ block: "start", inline: "nearest" });
        const focusTarget = focusSelector ? modal.querySelector(focusSelector) : null;
        focusTarget?.focus?.({ preventScroll: true });
      });
    }
  }

  function closeModalOverlay(overlay) {
    if (!overlay) return;
    overlay.classList.add("hidden");
    overlay.scrollTop = 0;
    const hasOpenOverlay = [els.modalOverlay, els.stageModalOverlay, els.historyModalOverlay, els.permissionModalOverlay]
      .some((item) => item && !item.classList.contains("hidden"));
    if (!hasOpenOverlay) document.body.classList.remove("modal-open");
  }

  function brMoney(value) {
    return Number(value || 0).toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
  }

  function formatPlanValue(value) {
    return Number(value || 0).toLocaleString("pt-BR", {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2
    });
  }

  function parseMonetaryValue(value) {
    if (value === null || value === undefined || value === "") return 0;

    let normalized = String(value)
      .trim()
      .replace(/r\$\s*/gi, "")
      .replace(/\s+/g, "")
      .replace(/[^0-9,.-]/g, "");

    if (!normalized) return 0;

    const lastComma = normalized.lastIndexOf(",");
    const lastDot = normalized.lastIndexOf(".");

    if (lastComma >= 0 && lastDot >= 0) {
      if (lastComma > lastDot) {
        normalized = normalized.replace(/\./g, "").replace(",", ".");
      } else {
        normalized = normalized.replace(/,/g, "");
      }
    } else if (lastComma >= 0) {
      normalized = normalized.replace(/\./g, "").replace(",", ".");
    } else if ((normalized.match(/\./g) || []).length > 1) {
      normalized = normalized.replace(/\./g, "");
    } else if (lastDot >= 0) {
      const decimalDigits = normalized.length - lastDot - 1;
      if (decimalDigits === 3) {
        normalized = normalized.replace(".", "");
      }
    }

    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function formatPlanInputValue(value) {
    if (value === null || value === undefined || value === "") return "";
    return formatPlanValue(parseMonetaryValue(value));
  }

  function normalizePlanNameKey(value) {
    return String(value || "").trim().toLowerCase();
  }

  function isNoPlanName(value) {
    return normalizePlanNameKey(value) === "sem plano";
  }

  function getPlansTotalValue(plans) {
    return cleanPlanList(plans).reduce((sum, item) => sum + Number(item.value || 0), 0);
  }

  function isMissingRelationError(error) {
    const code = String(error?.code || "").toUpperCase();
    const message = String(error?.message || "").toLowerCase();
    return code === "PGRST205" || code === "42P01" || message.includes("does not exist") || message.includes("could not find the table");
  }

  function isDuplicateKeyError(error) {
    const code = String(error?.code || "").toUpperCase();
    const message = String(error?.message || "").toLowerCase();
    return code === "23505" || message.includes("duplicate key") || message.includes("already exists");
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

  function getStoredHiddenPresetStageTypes() {
    try {
      const raw = window.localStorage.getItem("crmPax.hiddenPresetStageTypes");
      const parsed = JSON.parse(raw || "[]");
      return Array.isArray(parsed) ? parsed : [];
    } catch (_error) {
      return [];
    }
  }

  function persistHiddenPresetStageTypes(values) {
    const allowed = new Set(PRESET_STAGE_TYPES.map((item) => item.value));
    const normalized = [...new Set(
      (Array.isArray(values) ? values : [])
        .map((item) => String(item || "").trim())
        .filter((item) => allowed.has(item))
    )];

    try {
      window.localStorage.setItem("crmPax.hiddenPresetStageTypes", JSON.stringify(normalized));
    } catch (_error) {
      // Ignore storage failures and keep the in-memory list.
    }

    return normalized;
  }

  function runPeriodicStorageCleanup() {
    try {
      const now = Date.now();
      const lastCleanup = Number(window.localStorage.getItem(STORAGE_CLEANUP_KEY) || 0);

      if (!lastCleanup) {
        window.localStorage.setItem(STORAGE_CLEANUP_KEY, String(now));
        return;
      }

      if (now - lastCleanup < STORAGE_CLEANUP_INTERVAL_MS) return;

      STORAGE_CACHE_KEYS.forEach((key) => window.localStorage.removeItem(key));
      window.localStorage.setItem(STORAGE_CLEANUP_KEY, String(now));
    } catch (_error) {
      // Ignore storage failures and continue boot.
    }
  }

  function loadExternalScript(src) {
    return new Promise((resolve, reject) => {
      const normalizedSrc = String(src || "").trim();
      if (!ALLOWED_EXTERNAL_SCRIPT_URLS.has(normalizedSrc)) {
        reject(new Error(`Origem de script nao autorizada: ${normalizedSrc}`));
        return;
      }
      const script = document.createElement("script");
      script.src = normalizedSrc;
      script.async = true;
      script.crossOrigin = "anonymous";
      script.referrerPolicy = "no-referrer";
      script.onload = resolve;
      script.onerror = () => reject(new Error(`Falha ao carregar ${normalizedSrc}`));
      document.head.appendChild(script);
    });
  }

  async function ensureChartLibrary() {
    if (typeof window.Chart !== "undefined") return;
    if (!state.chartLoader) {
      state.chartLoader = loadExternalScript(CHART_JS_URL)
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

  function formatDayMonthLabel(value) {
    if (!value || !/^\d{4}-\d{2}-\d{2}$/.test(String(value))) return value || "-";
    return `${String(value).slice(8, 10)}/${String(value).slice(5, 7)}`;
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
        const value = parseMonetaryValue(item?.value);
        const contractNumber = String(item?.contract_number || item?.contractNumber || "").trim();
        const closedAt = normalizeDateInput(item?.closed_at || item?.closedAt || "") || "";

        return {
          name,
          value: Number.isFinite(value) ? value : 0,
          contract_number: contractNumber,
          closed_at: closedAt
        };
      })
      .filter((item) => item.name);
  }

  function planSupportsClosingDetails(plan = {}) {
    return !isNoPlanName(plan?.name) && Number(parseMonetaryValue(plan?.value)) > 0;
  }

  function getDefaultPlanName(index = 0) {
    const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    const safeIndex = Math.max(0, Number(index) || 0);
    const letter = alphabet[safeIndex % alphabet.length];
    const suffix = Math.floor(safeIndex / alphabet.length);
    return `Plano ${letter}${suffix ? suffix + 1 : ""}`;
  }

  function getPlanCatalog() {
    const catalog = new Map();

    state.leads.forEach((lead) => {
      getLeadPlans(lead).forEach((item) => {
        const name = String(item?.name || "").trim();
        if (!name || catalog.has(name)) return;
        catalog.set(name, Number(item?.value || 0));
      });
    });

    return [...catalog.entries()]
      .map(([name, value]) => ({ name, value: Number.isFinite(Number(value)) ? Number(value) : 0 }))
      .sort((a, b) => a.name.localeCompare(b.name, "pt-BR"));
  }

  function findKnownPlan(name) {
    const normalized = String(name || "").trim().toLowerCase();
    if (!normalized) return null;
    return getPlanCatalogWithDefault().find((item) => item.name.trim().toLowerCase() === normalized) || null;
  }

  function syncPlanDraftWithCatalog(index) {
    const draft = state.modalPlans[index];
    if (!draft) return false;

    const knownPlan = findKnownPlan(draft.name);
    if (!knownPlan) return false;

    draft.name = knownPlan.name;
    draft.value = knownPlan.value;
    return true;
  }

  function renderPlanSuggestions() {
    if (!els.planSuggestions) return;
    const catalog = getPlanCatalogWithDefault();
    els.planSuggestions.innerHTML = catalog
      .map((item) => `<option value="${escapeHtml(item.name)}"></option>`)
      .join("");
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
      } else if (!plans.length && Number(leadValue || 0) > 0) {
        plans.push({ name: getDefaultPlanName(0), value: Number(leadValue || 0) });
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
    const meta = getLeadMeta(lead?.notes || "", lead?.value || 0);
    const computedValue = getPlansTotalValue(meta.plans || []);
    return {
      ...lead,
      value: computedValue || Number(lead?.value || 0) || 0,
      start_date: normalizeDateInput(lead?.start_date || "") || String(lead?.start_date || "").trim(),
      _meta: meta
    };
  }

  function sanitizeHexColor(value, fallback = DEFAULT_STAGE_COLOR) {
    const raw = String(value || "").trim().toUpperCase();
    if (/^#[0-9A-F]{6}$/.test(raw)) return raw;

    const short = raw.match(/^#([0-9A-F]{3})$/);
    if (short) {
      return `#${short[1].split("").map((char) => char + char).join("")}`;
    }

    return fallback;
  }

  function normalizeStage(stage) {
    return {
      ...stage,
      color: sanitizeHexColor(stage?.color)
    };
  }

  function normalizeLeadSources(rows = []) {
    const map = new Map();

    DEFAULT_LEAD_SOURCES.forEach((name) => {
      const normalizedName = String(name || "").trim();
      if (!normalizedName) return;
      map.set(normalizedName.toLowerCase(), {
        id: normalizedName.toLowerCase(),
        name: normalizedName
      });
    });

    rows.forEach((item) => {
      const name = String(item?.name || "").trim();
      if (!name) return;
      map.set(name.toLowerCase(), {
        ...item,
        name
      });
    });

    return [...map.values()].sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pt-BR"));
  }

  function getLeadSourceNames() {
    return normalizeLeadSources(state.leadSources).map((item) => item.name);
  }

  function isSignupEmailAllowed(email) {
    if (!state.security.allowedSignupEmailDomains.length) return true;

    const normalizedEmail = String(email || "").trim().toLowerCase();
    const domain = normalizedEmail.split("@")[1] || "";
    return state.security.allowedSignupEmailDomains.includes(domain);
  }

  function getLeadPlans(lead) {
    return cleanPlanList(lead?._meta?.plans || []);
  }

  function getLeadPlan(lead) {
    const plans = getLeadPlans(lead);
    if (plans.length) return plans.map((item) => item.name).join(", ");
    return String(lead?._meta?.plan || "").trim();
  }

  function getLeadPlanDisplayText(lead) {
    const plans = getLeadPlans(lead)
      .map((item) => String(item?.name || "").trim())
      .filter((name) => name && !isNoPlanName(name));

    if (plans.length) return plans.join(", ");

    const rawPlan = String(lead?._meta?.plan || "").trim();
    if (rawPlan && !isNoPlanName(rawPlan)) return rawPlan;

    return "Não fechou ainda";
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
      lead?.traffic_type,
      lead?.social_source,
      lead?._meta?.legacyText,
      getLeadPlan(lead),
      ...getLeadPlans(lead).map((item) => `${item.name} ${item.value}`),
      ...getLeadObservations(lead).map((item) => item.text)
    ].join(" ").toLowerCase();
  }

  function getLeadMonthKey(lead) {
    const leadDate = getLeadDateKey(lead);
    return leadDate ? leadDate.slice(0, 7) : "";
  }

  function getLeadDateKey(lead) {
    const fromStart = normalizeDateInput(lead?.start_date || "");
    if (fromStart) return fromStart;

    return normalizeDateInput(String(lead?.created_at || "").slice(0, 10));
  }

  function getReportYear(leads) {
    const selectedMonth = String(els.monthFilter?.value || "").trim();
    if (/^\d{4}-\d{2}$/.test(selectedMonth)) return selectedMonth.slice(0, 4);

    const years = [...new Set(leads.map((lead) => getLeadDateKey(lead).slice(0, 4)).filter(Boolean))].sort();
    return years[years.length - 1] || String(new Date().getFullYear());
  }

  function buildYearDateKeys(year) {
    const yearNumber = Number(year);
    if (!Number.isInteger(yearNumber)) return [];

    const dates = [];
    const cursor = new Date(Date.UTC(yearNumber, 0, 1));

    while (cursor.getUTCFullYear() === yearNumber) {
      dates.push(cursor.toISOString().slice(0, 10));
      cursor.setUTCDate(cursor.getUTCDate() + 1);
    }

    return dates;
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
      { dropdown: els.monthFilterDropdown, btn: els.monthFilterBtn, menu: els.monthFilterMenu },
      { dropdown: els.leadSourceFilterDropdown, btn: els.leadSourceFilterBtn, menu: els.leadSourceFilterMenu }
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

  function isMobileViewport() {
    return window.matchMedia("(max-width: 700px)").matches;
  }

  function isCompactViewport() {
    return window.matchMedia("(max-width: 1100px)").matches;
  }

  function setMobileFiltersOpen(shouldOpen) {
    if (!els.mobileFiltersPanel || !els.mobileFiltersBtn) return;
    els.mobileFiltersPanel.classList.toggle("hidden", !shouldOpen);
    els.mobileFiltersBtn.classList.toggle("is-open", shouldOpen);
    els.mobileFiltersBtn.setAttribute("aria-expanded", shouldOpen ? "true" : "false");
    els.mobileFiltersBtn.setAttribute("aria-label", shouldOpen ? "Fechar filtros" : "Abrir filtros");
    requestAnimationFrame(() => {
      updateStickyLayout();
      syncPipelineScrollBars();
    });
  }

  function syncMobileFilterControls() {
    if (els.mobileOwnerFilter) els.mobileOwnerFilter.value = els.ownerFilter?.value || "";
    if (els.mobileMonthFilter) els.mobileMonthFilter.value = els.monthFilter?.value || "";
    if (els.mobileLeadSourceFilter) els.mobileLeadSourceFilter.value = els.leadSourceFilter?.value || "";
  }

  function normalizeMobileFilterTexts() {
    const topPanel = document.querySelector('#appScreen > #mobileFiltersPanel');
    topPanel?.querySelector('label[for="mobileOwnerFilter"]')?.replaceChildren("Responsável");
    topPanel?.querySelector('label[for="mobileMonthFilter"]')?.replaceChildren("Mês");

    const ownerDefault = topPanel?.querySelector('#mobileOwnerFilter option[value=""]');
    const monthDefault = topPanel?.querySelector('#mobileMonthFilter option[value=""]');
    topPanel?.querySelector('label[for="mobileLeadSourceFilter"]')?.replaceChildren("Origem do Lead");
    const leadSourceDefault = topPanel?.querySelector('#mobileLeadSourceFilter option[value=""]');
    if (ownerDefault) ownerDefault.textContent = "Todos os responsáveis";
    if (monthDefault) monthDefault.textContent = "Todos os meses";
    if (leadSourceDefault) leadSourceDefault.textContent = "Todas as origens";
  }

  function stageTypeLabel(type, customStageType = "") {
    const customLabel = String(customStageType || "").trim();
    if (type === "personalizado" && customLabel) return customLabel;

    const map = {
      andamento: "Andamento",
      fechado: "Fechado",
      cancelado: "Cancelado",
      espera: "Espera",
      personalizado: "Personalizado"
    };
    return map[type] || customLabel || "Andamento";
  }

  function updateStickyLayout() {
    const root = document.documentElement;
    if (!root) return;

    const mainRect = document.querySelector(".main")?.getBoundingClientRect?.();
    const mobileTopbarVisible = !!(els.mobileTopbar && window.getComputedStyle(els.mobileTopbar).display !== "none");
    const mobileTopbarHeight = mobileTopbarVisible ? els.mobileTopbar.offsetHeight : 0;
    const topbarHeight = els.topbar ? els.topbar.offsetHeight : 0;
    const metricsHeight = els.metricsSection ? els.metricsSection.offsetHeight : 0;
    const pipelineScrollbarHeight = els.pipelineScrollTop ? els.pipelineScrollTop.offsetHeight : 0;
    const funilStickyHeight = els.funilStickyHead ? els.funilStickyHead.offsetHeight : (metricsHeight + pipelineScrollbarHeight);
    const isFunilActive = state.activeView === "funil" && document.getElementById("view-funil")?.classList.contains("active-view");
    const pipelineViewportWidth = isFunilActive ? Math.max(0, Math.floor(mainRect?.width || 0)) : 0;
    const pipelineViewportLeft = isFunilActive ? Math.max(0, Math.floor(mainRect?.left || 0)) : 0;

    root.style.setProperty("--mobile-topbar-height", `${mobileTopbarHeight}px`);
    root.style.setProperty("--topbar-height", `${topbarHeight}px`);
    root.style.setProperty("--topbar-sticky-offset", `${mobileTopbarHeight}px`);
    root.style.setProperty("--metrics-sticky-offset", `${mobileTopbarHeight + topbarHeight}px`);
    root.style.setProperty("--pipeline-scrollbar-height", `${pipelineScrollbarHeight}px`);
    root.style.setProperty("--pipeline-scrollbar-sticky-offset", `${mobileTopbarHeight + topbarHeight + metricsHeight + 8}px`);
    root.style.setProperty("--funil-sticky-height", `${funilStickyHeight}px`);
    root.style.setProperty("--pipeline-scrollbar-fixed-left", `${pipelineViewportLeft}px`);
    root.style.setProperty("--pipeline-scrollbar-fixed-width", `${pipelineViewportWidth}px`);

    const topScrollbar = els.pipelineScrollTop;
    if (topScrollbar) {
      topScrollbar.style.left = `${pipelineViewportLeft}px`;
      topScrollbar.style.width = `${pipelineViewportWidth}px`;
      topScrollbar.style.bottom = "0px";
      topScrollbar.style.position = "fixed";
      topScrollbar.style.zIndex = "70";
    }
  }

  function bindOverlayDismiss(overlay, closeFn) {
    if (!overlay || typeof closeFn !== "function") return;

    let startedOnBackdrop = false;

    overlay.addEventListener("pointerdown", (event) => {
      startedOnBackdrop = event.target === overlay;
    });

    overlay.addEventListener("pointercancel", () => {
      startedOnBackdrop = false;
    });

    overlay.addEventListener("click", (event) => {
      if (startedOnBackdrop && event.target === overlay) {
        closeFn();
      }
      startedOnBackdrop = false;
    });
  }

  function isPresetStageType(value) {
    return PRESET_STAGE_TYPES.some((item) => item.value === value);
  }

  function getAvailablePresetStageTypes() {
    const hidden = new Set([
      ...state.hiddenPresetStageTypes,
      ...getStoredHiddenPresetStageTypes()
    ]);

    return PRESET_STAGE_TYPES.filter((item) => !hidden.has(item.value));
  }

  function getSelectableStageTypeOptions(includePersonalized = true) {
    const presetOptions = getAvailablePresetStageTypes().map((item) => ({
      value: item.value,
      label: item.label
    }));

    const customOptions = getCustomStageTypes().map((item) => ({
      value: `custom:${item}`,
      label: item
    }));

    const options = [...presetOptions, ...customOptions];
    if (includePersonalized) {
      options.push({ value: "personalizado", label: "Criar um novo" });
    }

    return options;
  }

  function getFallbackStageTypeSelection(excludedValue = "") {
    const nextOption = getSelectableStageTypeOptions(false).find((item) => item.value !== excludedValue);
    if (!nextOption) return { stage_type: "andamento", custom_stage_type: null };
    if (nextOption.value.startsWith("custom:")) {
      return {
        stage_type: "personalizado",
        custom_stage_type: nextOption.value.replace(/^custom:/, "")
      };
    }
    return { stage_type: nextOption.value, custom_stage_type: null };
  }

  function sortStageTypeNames(values) {
    return [...new Set(
      (Array.isArray(values) ? values : [])
        .map((item) => String(item || "").trim())
        .filter(Boolean)
    )].sort((a, b) => a.localeCompare(b, "pt-BR"));
  }

  function normalizeStageTypeNameKey(value) {
    return String(value || "").trim().toLowerCase();
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

  async function renameCustomStageType(oldName, nextName) {
    const previous = String(oldName || "").trim();
    const next = String(nextName || "").trim();
    if (!previous || !next || previous === next) return true;

    const { error: updateError } = await state.supabase
      .from("stages")
      .update({ custom_stage_type: next })
      .eq("custom_stage_type", previous);

    if (updateError) {
      console.error("Erro ao renomear tipo personalizado nas etapas:", updateError);
      return false;
    }

    const saved = await saveCustomStageType(next);
    if (!saved) return false;

    const removed = await deleteCustomStageTypeFromCatalog(previous);
    if (!removed) return false;

    state.customStageTypes = persistCustomStageTypes(
      state.customStageTypes.map((item) => (item === previous ? next : item))
    );

    return true;
  }

  function refreshStageTypeOptions(selectedValue = "", customValue = "") {
    if (!els.stageType) return;

    const finalOptions = getSelectableStageTypeOptions(true).map((item) => [item.value, item.label]);
    const fallbackValue = finalOptions[0]?.[0] || "personalizado";

    els.stageType.innerHTML = finalOptions.map(([value, label]) => `<option value="${escapeHtml(value)}">${escapeHtml(label)}</option>`).join("");

    const resolvedValue = selectedValue || (customValue ? `custom:${customValue}` : fallbackValue);
    els.stageType.value = finalOptions.some(([value]) => value === resolvedValue) ? resolvedValue : (customValue ? "personalizado" : fallbackValue);
    if (els.customStageType) els.customStageType.value = els.stageType.value === "personalizado" ? (customValue || "") : "";
    renderCustomStageTypeList();
    toggleCustomStageTypeField();
  }

  function renderCustomStageTypeList() {
    if (!els.savedStageTypes || !els.savedStageTypesGroup) return;

    const types = getSelectableStageTypeOptions(false);
    els.savedStageTypesGroup.classList.remove("hidden");

    if (!types.length) {
      els.savedStageTypes.innerHTML = '<div class="saved-stage-types-empty">Nenhum tipo disponivel na lista. Use "Criar um novo" para adicionar outro.</div>';
      return;
    }

    const selected = String(els.stageType?.value || "");
    els.savedStageTypes.innerHTML = types.map((type) => {
      const active = selected === type.value;
      return `
        <div class="saved-stage-type ${active ? "active" : ""}">
          <button type="button" class="saved-stage-type-select" data-stage-type-value="${escapeHtml(type.value)}">${escapeHtml(type.label)}</button>
        </div>
      `;
    }).join("");
  }

  function toggleCustomStageTypeField() {
    if (!els.customStageTypeGroup) return;
    const selected = String(els.stageType?.value || "andamento");
    const isNewCustom = selected === "personalizado";
    const isExistingCustom = selected.startsWith("custom:");
    const selectedOption = getSelectableStageTypeOptions(false).find((item) => item.value === selected);
    const selectedLabel = selectedOption?.label || stageTypeLabel(selected);

    const canEditCustom = isNewCustom || isExistingCustom;

    els.customStageTypeGroup.classList.remove("hidden");
    if (els.customStageType) {
      els.customStageType.disabled = !canEditCustom;
      if (isNewCustom) {
        els.customStageType.placeholder = "Criar um novo tipo";
      } else if (isExistingCustom) {
        els.customStageType.placeholder = "Personalize o tipo selecionado";
      } else {
        els.customStageType.placeholder = "Este tipo padrao nao pode ser personalizado aqui";
      }

      if (isNewCustom) {
        els.customStageType.value = els.customStageType.value || "";
      } else if (isExistingCustom) {
        els.customStageType.value = selected.replace(/^custom:/, "");
      } else {
        els.customStageType.value = selectedLabel;
      }
    }
    const canRemoveSelected = selected && selected !== "personalizado";
    if (els.savedStageTypeActions) els.savedStageTypeActions.classList.toggle("hidden", !canRemoveSelected);
    if (els.removeSelectedStageTypeBtn) {
      els.removeSelectedStageTypeBtn.classList.toggle("hidden", !canRemoveSelected);
      if (canRemoveSelected) {
        els.removeSelectedStageTypeBtn.textContent = `Remover tipo "${selectedOption?.label || stageTypeLabel(selected)}"`;
      }
    }
    if (els.removeCustomTypeBtn) els.removeCustomTypeBtn.classList.toggle("hidden", true);
    renderCustomStageTypeList();
  }

  async function removeCurrentCustomStageType() {
    const selected = String(els.stageType?.value || "");
    if (!selected || selected === "personalizado") {
      els.stageType.value = getSelectableStageTypeOptions(true)[0]?.value || "personalizado";
      if (els.customStageType) els.customStageType.value = "";
      toggleCustomStageTypeField();
      return;
    }

    const selectedOption = getSelectableStageTypeOptions(false).find((item) => item.value === selected);
    const selectedLabel = selectedOption?.label || stageTypeLabel(selected);
    if (!confirm(`Remover o tipo "${selectedLabel}" da lista? Os pipelines com esse tipo serao movidos para outro tipo disponivel.`)) return;

    let affected = [];
    if (selected.startsWith("custom:")) {
      const customValue = selected.replace(/^custom:/, "");
      if (!customValue) return;
      affected = state.stages.filter((stage) => String(stage.custom_stage_type || "").trim() === customValue);
    } else if (isPresetStageType(selected)) {
      affected = state.stages.filter((stage) => stage.stage_type === selected && !String(stage.custom_stage_type || "").trim());
    }

    const fallback = getFallbackStageTypeSelection(selected);
    if (affected.length) {
      const { error } = await state.supabase
        .from("stages")
        .update(fallback)
        .in("id", affected.map((stage) => stage.id));
      if (error) return alert(`Erro no Supabase: ${error.message}`);
    }

    if (selected.startsWith("custom:")) {
      const removed = await deleteCustomStageTypeFromCatalog(selected.replace(/^custom:/, ""));
      if (!removed) return alert("Nao foi possivel excluir o tipo personalizado da lista.");
    } else if (isPresetStageType(selected)) {
      state.hiddenPresetStageTypes = persistHiddenPresetStageTypes([
        ...state.hiddenPresetStageTypes,
        selected
      ]);
    }

    els.stageType.value = getSelectableStageTypeOptions(true).find((item) => item.value !== selected)?.value || "personalizado";
    if (els.customStageType) els.customStageType.value = "";
    toggleCustomStageTypeField();
    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  function getPlanCatalogWithDefault() {
    return [
      { name: "Sem plano", value: 0 },
      ...getPlanCatalog().filter((item) => String(item.name || "").trim().toLowerCase() !== "sem plano")
    ];
  }

  function getPlanSelectValue(name) {
    const normalized = String(name || "").trim().toLowerCase();
    if (!normalized) return "Sem plano";
    const knownPlan = getPlanCatalogWithDefault().find((item) => item.name.trim().toLowerCase() === normalized);
    return knownPlan ? knownPlan.name : "__custom__";
  }

  function applyPlanSelection(index, selectedName) {
    const draft = state.modalPlans[index];
    if (!draft) return;

    const normalized = String(selectedName || "").trim();
    if (!normalized || normalized === "__custom__") {
      if (!draft.name) draft.name = getDefaultPlanName(index);
      return;
    }

    const knownPlan = getPlanCatalogWithDefault().find((item) => item.name === normalized);
    if (!knownPlan) return;

    draft.name = knownPlan.name;
    draft.value = Number(knownPlan.value || 0);
  }

  function getLeadPlanValueText(lead) {
    const plans = getLeadPlans(lead).filter((item) => !isNoPlanName(item.name));
    if (!plans.length) {
      const rawValue = Number(lead?.value || 0);
      return rawValue > 0 ? formatPlanValue(rawValue) : "Não fechou ainda";
    }

    return plans.map((item) => `${item.name}: ${formatPlanValue(item.value || 0)}`).join(" | ");
  }

  async function syncPlanValuesAcrossLeads(referencePlans, excludeLeadId = null) {
    const targetPlans = cleanPlanList(referencePlans).filter((item) => !isNoPlanName(item.name));
    if (!targetPlans.length) return 0;

    const targetMap = new Map(
      targetPlans.map((item) => [normalizePlanNameKey(item.name), Number(item.value || 0)])
    );

    const pendingUpdates = state.leads.reduce((acc, lead) => {
      if (lead.id === excludeLeadId) return acc;

      const currentPlans = getLeadPlans(lead);
      if (!currentPlans.length) return acc;

      let changed = false;
      const nextPlans = currentPlans.map((item) => {
        const key = normalizePlanNameKey(item.name);
        if (!targetMap.has(key)) return item;

        const nextValue = targetMap.get(key);
        if (Number(item.value || 0) === nextValue) return item;

        changed = true;
        return { ...item, value: nextValue };
      });

      if (!changed) return acc;

      const leadMeta = getLeadMeta(lead?.notes || "", lead?.value || 0);
      acc.push({
        id: lead.id,
        name: lead.name,
        payload: {
          value: getPlansTotalValue(nextPlans),
          notes: serializeLeadMeta({
            ...leadMeta,
            plans: nextPlans,
            plan: nextPlans[0]?.name || leadMeta.plan || "",
            legacyText: leadMeta.legacyText,
            observations: leadMeta.observations
          })
        }
      });
      return acc;
    }, []);

    if (!pendingUpdates.length) return 0;

    const results = await Promise.all(
      pendingUpdates.map((item) =>
        state.supabase
          .from("leads")
          .update(item.payload)
          .eq("id", item.id)
      )
    );

    const failed = results.find((result) => result.error);
    if (failed?.error) throw failed.error;

    await logChange(
      "sync",
      "lead",
      excludeLeadId,
      `${pendingUpdates.length} lead(s) tiveram o valor do plano sincronizado por ${getUserDisplayName()}.`,
      {
        plans: targetPlans,
        affected_leads: pendingUpdates.map((item) => ({ id: item.id, name: item.name }))
      }
    );

    return pendingUpdates.length;
  }

  function renderPlanItems() {
    if (!els.plansList) return;
    renderPlanSuggestions();

    if (!state.modalPlans.length) {
      els.plansList.innerHTML = '<div class="plan-empty">Nenhum plano adicionado ainda.</div>';
      return;
    }

    els.plansList.innerHTML = state.modalPlans.map((item, index) => `
      <div class="plan-item" data-index="${index}">
        <div class="plan-item-grid">
          <select class="plan-select-input">
            ${getPlanCatalogWithDefault().map((plan) => `<option value="${escapeHtml(plan.name)}" ${getPlanSelectValue(item.name) === plan.name ? "selected" : ""}>${escapeHtml(plan.name)}</option>`).join("")}
            <option value="__custom__" ${getPlanSelectValue(item.name) === "__custom__" ? "selected" : ""}>Personalizar nome</option>
          </select>
          <input type="text" class="plan-name-input" placeholder="Nome do plano" list="planSuggestions" value="${escapeHtml(item.name || "")}" />
          <input type="text" class="plan-value-input" placeholder="Valor do plano" inputmode="decimal" value="${escapeHtml(formatPlanInputValue(item.value))}" />
        </div>
        ${planSupportsClosingDetails(item) ? `
          <div class="plan-item-extra">
            <div class="plan-item-extra-title">Dados do fechamento</div>
            <div class="plan-item-extra-grid">
              <input type="text" class="plan-contract-input" placeholder="Numero do contrato" value="${escapeHtml(item.contract_number || "")}" />
              <input type="date" class="plan-closed-date-input" value="${escapeHtml(item.closed_at || "")}" />
            </div>
          </div>
        ` : ""}
        <div class="plan-item-actions">
          <button type="button" class="btn btn-danger plan-remove-btn">Excluir</button>
        </div>
      </div>
    `).join("");
  }

  function addPlanFromDraft() {
    state.modalPlans.push({ name: "Sem plano", value: 0 });
    applyPlanSelection(state.modalPlans.length - 1, "Sem plano");
    renderPlanItems();
  }

  function setupPlanListEvents() {
    const handlePlanDraftInput = (event) => {
      const item = event.target.closest(".plan-item");
      if (!item) return;
      const index = Number(item.dataset.index);
      if (Number.isNaN(index) || !state.modalPlans[index]) return;
      if (event.target.classList.contains("plan-select-input")) {
        applyPlanSelection(index, event.target.value);
        renderPlanItems();
        return;
      }
      if (event.target.classList.contains("plan-name-input")) {
        const hadClosingDetails = planSupportsClosingDetails(state.modalPlans[index]);
        state.modalPlans[index].name = event.target.value;
        if (syncPlanDraftWithCatalog(index) || hadClosingDetails !== planSupportsClosingDetails(state.modalPlans[index])) {
          renderPlanItems();
        }
        return;
      }
      if (event.target.classList.contains("plan-value-input")) {
        const hadClosingDetails = planSupportsClosingDetails(state.modalPlans[index]);
        state.modalPlans[index].value = event.target.value;
        if (hadClosingDetails !== planSupportsClosingDetails(state.modalPlans[index])) {
          renderPlanItems();
          return;
        }
        return;
      }
      if (event.target.classList.contains("plan-contract-input")) {
        state.modalPlans[index].contract_number = event.target.value;
        return;
      }
      if (event.target.classList.contains("plan-closed-date-input")) {
        state.modalPlans[index].closed_at = normalizeDateInput(event.target.value);
      }
    };

    els.plansList?.addEventListener("input", handlePlanDraftInput);
    els.plansList?.addEventListener("change", handlePlanDraftInput);

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

  function getPipelineScrollMetrics() {
    const area = els.pipelineScrollArea;
    const top = els.pipelineScrollTop;
    const isFunilActive = state.activeView === "funil" && document.getElementById("view-funil")?.classList.contains("active-view");
    const contentWidth = Math.max(
      area?.scrollWidth || 0,
      els.pipeline?.scrollWidth || 0
    );
    const viewportWidth = Math.floor(area?.getBoundingClientRect().width || area?.clientWidth || 0);
    const maxScrollLeft = Math.max(0, contentWidth - viewportWidth);
    const hasHorizontalOverflow = maxScrollLeft > 4;
    const shouldShowTop = hasHorizontalOverflow && isFunilActive;
    const trackWidth = Math.floor(els.pipelineScrollTopTrack?.getBoundingClientRect().width || top?.clientWidth || 0);

    return {
      area,
      top,
      contentWidth,
      viewportWidth,
      maxScrollLeft,
      hasHorizontalOverflow,
      shouldShowTop,
      trackWidth
    };
  }

  function updatePipelineScrollThumb(metrics = getPipelineScrollMetrics()) {
    const thumb = els.pipelineScrollTopThumb;
    if (!thumb) return;

    if (!metrics.shouldShowTop || !metrics.trackWidth || !metrics.viewportWidth || metrics.contentWidth <= metrics.viewportWidth) {
      thumb.style.width = "0px";
      thumb.style.transform = "translate(0, -50%)";
      return;
    }

    const area = metrics.area;
    const scrollLeft = Math.min(metrics.maxScrollLeft, Math.max(0, area?.scrollLeft || 0));
    const thumbWidth = Math.max(48, Math.round((metrics.viewportWidth / metrics.contentWidth) * metrics.trackWidth));
    const maxThumbOffset = Math.max(0, metrics.trackWidth - thumbWidth);
    const thumbOffset = metrics.maxScrollLeft > 0
      ? Math.round((scrollLeft / metrics.maxScrollLeft) * maxThumbOffset)
      : 0;

    thumb.style.width = `${thumbWidth}px`;
    thumb.style.transform = `translate(${thumbOffset}px, -50%)`;
  }

  function syncPipelineScrollBars(source = null) {
    let metrics = getPipelineScrollMetrics();
    const area = metrics.area;
    const top = metrics.top;

    if (top) top.style.display = metrics.shouldShowTop ? "flex" : "none";
    if (!area) return;

    if (typeof source === "number") {
      area.scrollLeft = Math.min(metrics.maxScrollLeft, Math.max(0, source));
    }

    // The custom scrollbar track has no measurable width while hidden.
    // Re-read metrics after making the bar visible so the thumb can be sized.
    if (metrics.shouldShowTop && !metrics.trackWidth) {
      metrics = getPipelineScrollMetrics();
    }

    updatePipelineScrollThumb(metrics);
  }

  function setPipelineScrollFromPointer(clientX) {
    const metrics = getPipelineScrollMetrics();
    const track = els.pipelineScrollTopTrack;
    const thumb = els.pipelineScrollTopThumb;
    if (!metrics.area || !track || !thumb || !metrics.maxScrollLeft || !metrics.trackWidth) return;

    const rect = track.getBoundingClientRect();
    const thumbWidth = Math.max(thumb.offsetWidth || 0, 48);
    const maxThumbOffset = Math.max(0, rect.width - thumbWidth);
    const targetOffset = Math.min(
      maxThumbOffset,
      Math.max(0, (clientX - rect.left) - (thumbWidth / 2))
    );
    const nextScrollLeft = maxThumbOffset > 0
      ? (targetOffset / maxThumbOffset) * metrics.maxScrollLeft
      : 0;

    syncPipelineScrollBars(nextScrollLeft);
  }

  function startPipelineScrollbarDrag(event) {
    if (event.button !== 0) return;

    const metrics = getPipelineScrollMetrics();
    if (!metrics.shouldShowTop || !metrics.maxScrollLeft) return;

    state.pipelineScrollbarDrag = {
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startScrollLeft: metrics.area?.scrollLeft || 0
    };

    els.pipelineScrollTopThumb?.setPointerCapture?.(event.pointerId);
    event.preventDefault();
  }

  function stopPipelineScrollbarDrag(event) {
    const drag = state.pipelineScrollbarDrag;
    if (!drag) return;
    if (event?.pointerId !== undefined && drag.pointerId !== event.pointerId) return;

    try {
      els.pipelineScrollTopThumb?.releasePointerCapture?.(drag.pointerId);
    } catch (_error) {
      // Ignore release failures when capture was already lost.
    }

    state.pipelineScrollbarDrag = null;
  }

  function handlePipelineScrollbarDrag(event) {
    const drag = state.pipelineScrollbarDrag;
    if (!drag || drag.pointerId !== event.pointerId) return;

    const metrics = getPipelineScrollMetrics();
    const track = els.pipelineScrollTopTrack;
    const thumb = els.pipelineScrollTopThumb;
    if (!metrics.maxScrollLeft || !track || !thumb) return;

    const thumbWidth = Math.max(thumb.offsetWidth || 0, 48);
    const maxThumbOffset = Math.max(1, track.clientWidth - thumbWidth);
    const deltaX = event.clientX - drag.startClientX;
    const scrollDelta = (deltaX / maxThumbOffset) * metrics.maxScrollLeft;

    syncPipelineScrollBars(drag.startScrollLeft + scrollDelta);
    event.preventDefault();
  }

  function attachPipelineScrollEvents() {
    const area = els.pipelineScrollArea;
    const track = els.pipelineScrollTopTrack;
    const thumb = els.pipelineScrollTopThumb;
    if (!area) return;

    area.addEventListener("scroll", () => syncPipelineScrollBars());
    window.addEventListener("resize", () => syncPipelineScrollBars());
    track?.addEventListener("pointerdown", (event) => {
      if (event.target === thumb) return;
      if (event.button !== 0) return;
      setPipelineScrollFromPointer(event.clientX);
      event.preventDefault();
    });
    thumb?.addEventListener("pointerdown", startPipelineScrollbarDrag);
    thumb?.addEventListener("pointermove", handlePipelineScrollbarDrag);
    thumb?.addEventListener("pointerup", stopPipelineScrollbarDrag);
    thumb?.addEventListener("pointercancel", stopPipelineScrollbarDrag);

    if (typeof ResizeObserver === "function" && !state.pipelineScrollObserver) {
      state.pipelineScrollObserver = new ResizeObserver(() => {
        syncPipelineScrollBars();
        requestAnimationFrame(updateStickyLayout);
      });
      state.pipelineScrollObserver.observe(area);
      if (els.pipeline) state.pipelineScrollObserver.observe(els.pipeline);
    }
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
    const color = sanitizeHexColor(value);
    if (els.stageColor && els.stageColor.value !== color) {
      els.stageColor.value = color;
    }
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
    if (!canManageStages()) {
      alert("Somente administradores podem reordenar pipelines.");
      return;
    }

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
      const accessAllowed = await enforceApprovedSession();
      if (!accessAllowed) return;
      await enterApp();
      return;
    }

    showScreen("authScreen");
  }

  async function handleAuthStateChange(event, session) {
    if (event === "PASSWORD_RECOVERY") {
      els.resetPasswordBox.classList.remove("hidden");
      showScreen("authScreen");
      setMessage(els.authMessage, "Digite sua nova senha e salve.");
      return;
    }

    if (event === "SIGNED_OUT") {
      resetAppState();
      showScreen("authScreen");
      return;
    }

    if (event !== "SIGNED_IN" || !session?.user) return;

    state.currentUser = session.user;
    await ensureProfile();
    const accessAllowed = await enforceApprovedSession();
    if (!accessAllowed) return;
    if (els.appScreen.classList.contains("hidden")) {
      await enterApp();
    } else {
      await loadAppData({ includeProfiles: state.profilesLoaded });
    }
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

    state.profile = {
      ...payload,
      ...profile,
      role: profile?.role || USER_ROLE.USER,
      access_status: profile?.access_status || ACCESS_STATUS.PENDING
    };
    applyRoleBasedUi();
  }

  async function enforceApprovedSession() {
    if (!state.currentUser) return false;

    const status = getAccessStatus();
    if (status === ACCESS_STATUS.APPROVED) return true;

    if (status === ACCESS_STATUS.REJECTED) {
      await forceSignOutWithMessage("Seu acesso foi recusado. Solicite nova analise com o administrador.");
      return false;
    }

    await forceSignOutWithMessage("Seu acesso ainda esta pendente de aprovacao administrativa.");
    return false;
  }

  async function enterApp() {
    showScreen("appScreen");
    els.orgNameLabel.textContent = "CRM Pax";
    els.mobileOrgName.textContent = "CRM Pax";
    els.userWelcome.textContent = getUserDisplayName();
    applyRoleBasedUi();

    await loadAppData({ includeProfiles: true });
    bindView("funil");
  }

  async function loadAppData(options = {}) {
    const includeProfiles = options.includeProfiles === true;
    const includeAdminData = canManageAdminAreas();
    const [stagesRes, leadsRes, profilesRes, stageTypesRes, leadSourcesRes, accessRequestsRes, adminRequestsRes] = await Promise.all([
      state.supabase.from("stages").select("*").order("position", { ascending: true }),
      state.supabase.from("leads").select("*").order("created_at", { ascending: false }),
      includeProfiles
        ? state.supabase.from("profiles").select("*").order("full_name", { ascending: true })
        : Promise.resolve({ data: state.profiles, error: null }),
      state.supabase.from("stage_type_catalog").select("name").order("name", { ascending: true }),
      state.supabase.from("lead_source_catalog").select("*").order("name", { ascending: true }),
      includeAdminData
        ? state.supabase.from("access_requests").select("*").order("created_at", { ascending: false })
        : Promise.resolve({ data: [], error: null }),
      includeAdminData
        ? state.supabase.from("admin_requests").select("*").order("created_at", { ascending: false })
        : Promise.resolve({ data: [], error: null })
    ]);

    if (stagesRes.error) console.error(stagesRes.error);
    if (leadsRes.error) console.error(leadsRes.error);
    if (includeProfiles && profilesRes.error) console.error(profilesRes.error);
    if (stageTypesRes.error && !isMissingRelationError(stageTypesRes.error)) console.error(stageTypesRes.error);
    if (leadSourcesRes.error && !isMissingRelationError(leadSourcesRes.error)) console.error(leadSourcesRes.error);
    if (includeAdminData && accessRequestsRes.error && !isMissingRelationError(accessRequestsRes.error)) console.error(accessRequestsRes.error);
    if (includeAdminData && adminRequestsRes.error && !isMissingRelationError(adminRequestsRes.error)) console.error(adminRequestsRes.error);

    const firstError =
      stagesRes.error ||
      leadsRes.error ||
      profilesRes.error ||
      (stageTypesRes.error && !isMissingRelationError(stageTypesRes.error) ? stageTypesRes.error : null) ||
      (leadSourcesRes.error && !isMissingRelationError(leadSourcesRes.error) ? leadSourcesRes.error : null) ||
      (accessRequestsRes.error && !isMissingRelationError(accessRequestsRes.error) ? accessRequestsRes.error : null) ||
      (adminRequestsRes.error && !isMissingRelationError(adminRequestsRes.error) ? adminRequestsRes.error : null);
    if (firstError) {
      alert(`Erro ao carregar dados do Supabase: ${firstError.message}`);
      return;
    }

    state.stages = (stagesRes.data || []).map(normalizeStage);
    state.leads = (leadsRes.data || []).map(normalizeLead);
    state.leadSources = normalizeLeadSources(leadSourcesRes.data || []);
    state.customStageTypes = persistCustomStageTypes([
      ...getStoredCustomStageTypes(),
      ...(stageTypesRes.data || []).map((item) => item?.name)
    ]);
    state.hiddenPresetStageTypes = persistHiddenPresetStageTypes(getStoredHiddenPresetStageTypes());
    if (includeProfiles) {
      state.profiles = profilesRes.data || [];
      state.profilesLoaded = true;
    }
    state.accessRequests = accessRequestsRes.data || [];
    state.adminRequests = adminRequestsRes.data || [];
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
    if (!canViewHistory()) {
      state.history = [];
      state.historyLoaded = true;
      renderHistoryText();
      return;
    }

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
    const leadSource = els.leadSourceFilter?.value || "";

    return state.leads.filter((lead) => {
      const matchesSearch = !search || getLeadSearchText(lead).includes(search);
      const matchesOwner = !owner || lead.owner === owner;
      const matchesMonth = !month || getLeadMonthKey(lead) === month;
      const matchesLeadSource = !leadSource || String(lead.traffic_type || "") === leadSource;

      return matchesSearch && matchesOwner && matchesMonth && matchesLeadSource;
    });
  }

  function populateFilters() {
    const owners = [...new Set(state.leads.map((x) => x.owner).filter(Boolean))].sort();
    const months = [...new Set(state.leads.map((lead) => getLeadMonthKey(lead)).filter(Boolean))].sort();
    const leadSources = getLeadSourceNames();

    const currentOwner = els.ownerFilter.value;
    const currentMonth = els.monthFilter.value;
    const currentLeadSource = els.leadSourceFilter?.value || "";

    els.ownerFilter.innerHTML =
      '<option value="">Todos os responsáveis</option>' +
      owners.map((o) => `<option value="${escapeHtml(o)}">${escapeHtml(o)}</option>`).join("");

    els.monthFilter.innerHTML =
      '<option value="">Todos os meses</option>' +
      months.map((m) => `<option value="${m}">${formatMonthLabel(m)}</option>`).join("");

    if (els.leadSourceFilter) {
      els.leadSourceFilter.innerHTML =
        '<option value="">Todas as origens</option>' +
        leadSources.map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`).join("");
    }

    if (els.mobileOwnerFilter) {
      els.mobileOwnerFilter.innerHTML =
        '<option value="">Todos os responsáveis</option>' +
        owners.map((o) => `<option value="${escapeHtml(o)}">${escapeHtml(o)}</option>`).join("");
    }

    if (els.mobileMonthFilter) {
      els.mobileMonthFilter.innerHTML =
        '<option value="">Todos os meses</option>' +
        months.map((m) => `<option value="${m}">${formatMonthLabel(m)}</option>`).join("");
    }

    if (els.mobileLeadSourceFilter) {
      els.mobileLeadSourceFilter.innerHTML =
        '<option value="">Todas as origens</option>' +
        leadSources.map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`).join("");
    }

    els.stage.innerHTML = state.stages
      .map((s) => `<option value="${s.id}">${escapeHtml(s.name)}</option>`)
      .join("");

    if (els.trafficType) {
      els.trafficType.innerHTML =
        '<option value="">Selecione</option>' +
        leadSources.map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`).join("");
    }

    refreshStageTypeOptions(els.stageType?.value || "andamento");

    els.ownerFilter.value = owners.includes(currentOwner) ? currentOwner : "";
    els.monthFilter.value = months.includes(currentMonth) ? currentMonth : "";
    if (els.leadSourceFilter) {
      els.leadSourceFilter.value = leadSources.includes(currentLeadSource) ? currentLeadSource : "";
    }
    syncMobileFilterControls();

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

    renderFilterOptions(
      els.leadSourceFilter,
      els.leadSourceFilterMenu,
      els.leadSourceFilterLabel,
      "Todas as origens",
      (_value, text) => text
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

    if (els.reportTotalLeads) els.reportTotalLeads.textContent = metrics.total;
    els.reportClosedValue.textContent = brMoney(metrics.totalValue);
    if (els.reportClosedDeals) els.reportClosedDeals.textContent = metrics.closed;
    els.reportPaidCount.textContent = metrics.paidCount;
    els.reportOrganicCount.textContent = metrics.organicCount;
    els.reportAvgTicket.textContent = brMoney(metrics.avgTicket);
    if (els.reportConversionRate) els.reportConversionRate.textContent = `${metrics.conversion.toFixed(1)}%`;
    els.reportTopOwner.textContent = metrics.topOwner?.[0] || "-";
    els.reportTopStage.textContent = metrics.topStage?.name || "-";
    els.reportBestMonth.textContent = metrics.bestMonth?.[0] ? formatMonthLabel(metrics.bestMonth[0]) : "-";
    if (els.reportClosedPlans) els.reportClosedPlans.textContent = metrics.planSummary.reduce((sum, item) => sum + item.count, 0);

    if (els.planSummaryBody) {
      els.planSummaryBody.innerHTML = metrics.planSummary.length
        ? metrics.planSummary.map((item) => `
          <tr>
            <td>${escapeHtml(item.plan)}</td>
            <td>${formatPlanValue(item.unitValue)}</td>
            <td>${item.count}</td>
            <td>${formatPlanValue(item.totalValue)}</td>
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
                <div class="card-value">${escapeHtml(getLeadPlanValueText(lead))}</div>
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
              <span><strong>Canal de origem:</strong> ${escapeHtml(lead.social_source || "-")}</span>
            </div>

            ${getLeadLatestObservation(lead) ? `<div class="card-notes"><strong>Ultima observacao:</strong> ${escapeHtml(getLeadLatestObservation(lead).text)}${getLeadLatestObservation(lead).date ? `<small>${formatDate(getLeadLatestObservation(lead).date)}</small>` : ""}</div>` : ""}

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
              ${canManageStages() ? `
                <div class="column-order-actions">
                  <button type="button" class="icon-btn" data-stage-action="move-left" data-id="${stage.id}" ${index === 0 ? "disabled" : ""} title="Mover para a esquerda">&larr;</button>
                  <button type="button" class="icon-btn" data-stage-action="move-right" data-id="${stage.id}" ${index === state.stages.length - 1 ? "disabled" : ""} title="Mover para a direita">&rarr;</button>
                </div>
              ` : ""}
            </div>
            <span class="badge">${leads.length}</span>
          </header>
          <div class="column-body">${cards}</div>
        </section>
      `;
    }).join("");

    bindPipelineEvents();
    requestAnimationFrame(() => {
      syncPipelineScrollBars();
      updateStickyLayout();
    });
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
    if (!canDeleteLeads()) {
      requestAdminAuthorization({
        requestType: "bulk_delete_leads",
        title: "Solicitar exclusao em lote",
        description: "Voce nao tem permissao para excluir leads em lote. Envie uma solicitacao para o administrador.",
        entityType: "lead",
        entityId: null,
        payload: {
          lead_ids: [...state.selectedLeadIds],
          lead_names: state.leads.filter((lead) => state.selectedLeadIds.has(lead.id)).map((lead) => lead.name)
        }
      });
      return;
    }

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
          <td colspan="11" class="empty-state">Nenhum lead encontrado.</td>
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
        <td>${escapeHtml(getLeadPlanValueText(lead))}</td>
        <td>${formatDate(lead.start_date)}</td>
        <td>${escapeHtml(lead.traffic_type || "-")}</td>
        <td>${escapeHtml(lead.social_source || "-")}</td>
        <td>${escapeHtml(getLeadPlanDisplayText(lead))}</td>
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

    const people = (canManageAdminAreas()
      ? state.profiles
      : state.profiles.filter((p) => activeIds.has(p.id)))
      .sort((a, b) => String(a.full_name || "").localeCompare(String(b.full_name || "")));

    els.teamList.innerHTML = people.map((profile) => `
      <div class="team-item">
        <div class="team-item-head">
          <strong>${escapeHtml(profile.full_name || "Usuario")}</strong>
          <span class="status-chip ${escapeHtml(String(profile.access_status || ACCESS_STATUS.PENDING).toLowerCase())}">${escapeHtml(getAccessStatusLabel(String(profile.access_status || ACCESS_STATUS.PENDING).toLowerCase()))}</span>
        </div>
        <div class="team-item-meta">E-mail: ${escapeHtml(profile.email || "-")}</div>
        <div class="team-item-meta">Nivel: ${escapeHtml(getRoleLabel(String(profile.role || USER_ROLE.USER).toLowerCase()))}</div>
        ${canManageAdminAreas() ? `
          <div class="team-item-actions">
            <select class="team-role-select" data-profile-id="${profile.id}">
              <option value="${USER_ROLE.USER}" ${String(profile.role || "").toLowerCase() === USER_ROLE.USER ? "selected" : ""}>Usuario comum</option>
              <option value="${USER_ROLE.ADMIN}" ${String(profile.role || "").toLowerCase() === USER_ROLE.ADMIN ? "selected" : ""}>Administrador</option>
            </select>
            <button
              type="button"
              class="delete-btn"
              data-team-action="delete"
              data-profile-id="${profile.id}"
              ${state.currentUser?.id === profile.id ? "disabled" : ""}
            >
              Excluir
            </button>
          </div>
        ` : ""}
      </div>
    `).join("") || '<div class="team-item">Nenhum usuario encontrado.</div>';
  }

  function renderStagesConfig() {
    if (!canManageStages()) {
      els.stagesConfigList.innerHTML = '<div class="stage-config-item">Somente administradores podem gerenciar pipelines.</div>';
      return;
    }

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

  function renderLeadSourcesConfig() {
    if (!els.leadSourcesConfigList) return;

    if (!canManageLeadSources()) {
      els.leadSourcesConfigList.innerHTML = '<div class="saved-stage-types-empty">Somente administradores podem gerenciar origens do lead.</div>';
      return;
    }

    const items = getLeadSourceNames();
    if (!items.length) {
      els.leadSourcesConfigList.innerHTML = '<div class="saved-stage-types-empty">Nenhuma origem cadastrada.</div>';
      return;
    }

    els.leadSourcesConfigList.innerHTML = items.map((name) => `
      <div class="saved-stage-type">
        <button type="button" class="saved-stage-type-select" data-source-action="edit" data-source-name="${escapeHtml(name)}">${escapeHtml(name)}</button>
        <button type="button" class="saved-stage-type-delete" data-source-action="delete" data-source-name="${escapeHtml(name)}">Excluir</button>
      </div>
    `).join("");
  }

  function renderRequests() {
    if (!els.accessRequestsList || !els.adminRequestsList) return;

    if (!canManageAdminAreas()) {
      els.accessRequestsList.innerHTML = "";
      els.adminRequestsList.innerHTML = "";
      return;
    }

    const pendingAccessRequests = state.accessRequests.filter((request) => String(request.status || ACCESS_STATUS.PENDING).toLowerCase() === ACCESS_STATUS.PENDING);
    const pendingAdminRequests = state.adminRequests.filter((request) => String(request.status || ACCESS_STATUS.PENDING).toLowerCase() === ACCESS_STATUS.PENDING);

    const accessCards = pendingAccessRequests.length
      ? pendingAccessRequests.map((request) => `
        <article class="request-card">
          <div class="request-card-head">
            <div>
              <strong>${escapeHtml(request.full_name || request.email || "Solicitacao de acesso")}</strong>
              <div class="request-card-meta">${escapeHtml(request.email || "-")}</div>
            </div>
            <span class="status-chip ${escapeHtml(String(request.status || ACCESS_STATUS.PENDING).toLowerCase())}">${escapeHtml(getAccessStatusLabel(String(request.status || ACCESS_STATUS.PENDING).toLowerCase()))}</span>
          </div>
          <div class="request-card-meta">Solicitado em: ${request.created_at ? new Date(request.created_at).toLocaleString("pt-BR") : "-"}</div>
          <div class="request-card-actions">
            <select data-access-role="${request.id}">
              <option value="${USER_ROLE.USER}">Usuario comum</option>
              <option value="${USER_ROLE.ADMIN}">Administrador</option>
            </select>
            <button type="button" class="btn btn-primary" data-access-action="approve" data-id="${request.id}">Aprovar</button>
            <button type="button" class="btn btn-secondary" data-access-action="reject" data-id="${request.id}">Recusar</button>
          </div>
        </article>
      `).join("")
      : '<div class="request-card request-card-empty">Nenhuma solicitacao de acesso pendente.</div>';

    const adminCards = pendingAdminRequests.length
      ? pendingAdminRequests.map((request) => `
        <article class="request-card">
          <div class="request-card-head">
            <div>
              <strong>${escapeHtml(request.title || request.request_type || "Solicitacao operacional")}</strong>
              <div class="request-card-meta">${escapeHtml(request.requested_by_name || request.requested_by_email || "-")}</div>
            </div>
            <span class="status-chip ${escapeHtml(String(request.status || ACCESS_STATUS.PENDING).toLowerCase())}">${escapeHtml(getAccessStatusLabel(String(request.status || ACCESS_STATUS.PENDING).toLowerCase()))}</span>
          </div>
          <div class="request-card-meta">${escapeHtml(request.description || "")}</div>
          ${request.reason ? `<div class="request-card-reason">${escapeHtml(request.reason)}</div>` : ""}
          <div class="request-card-actions">
            <button type="button" class="btn btn-primary" data-admin-request-action="approve" data-id="${request.id}">Aprovar</button>
            <button type="button" class="btn btn-secondary" data-admin-request-action="reject" data-id="${request.id}">Recusar</button>
          </div>
        </article>
      `).join("")
      : '<div class="request-card request-card-empty">Nenhuma solicitacao operacional registrada.</div>';

    els.accessRequestsList.innerHTML = accessCards;
    els.adminRequestsList.innerHTML = adminCards;
  }

  function renderHistoryText() {
    if (!canViewHistory()) {
      els.historyText.textContent = "Historico disponivel apenas para administradores.";
      return;
    }
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
    return state.stages.map((stage) => sanitizeHexColor(stage.color));
  }

  function renderCharts() {
    if (typeof Chart === "undefined") return;

    ["pipeline", "traffic", "owner", "yearlyDaily", "monthly", "social", "planCount", "planRevenue"].forEach(destroyChart);

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

    const reportYear = getReportYear(filteredAllMonths);
    const yearDateKeys = buildYearDateKeys(reportYear);
    const yearDayMap = filteredAllMonths.reduce((acc, lead) => {
      const key = getLeadDateKey(lead);
      if (!key || !key.startsWith(`${reportYear}-`)) return acc;
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});

    const socialMap = filtered.reduce((acc, lead) => {
      const key = lead.social_source || "Não informado";
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});

    const planLabels = metrics.planSummary.map((item) => `${item.plan} - ${formatPlanValue(item.unitValue)}`);
    const yearChartTitle = $("yearlyChartTitle");
    if (yearChartTitle) yearChartTitle.textContent = `Contatos diários de ${reportYear} (dia a dia)`;

    const create = (key, id, config) => {
      const canvas = $(id);
      if (!canvas) return;
      state.charts[key] = new Chart(canvas, config);
    };

    const yearLabels = yearDateKeys.map((key) => formatDayMonthLabel(key));
    const monthStartIndexes = new Set(yearDateKeys.reduce((acc, key, index) => {
      if (key.endsWith("-01")) acc.push(index);
      return acc;
    }, []));

    create("pipeline", "pipelineChart", makeChartConfig("bar", metrics.byStage.map((item) => item.name), [{ label: "Leads", data: metrics.byStage.map((item) => item.count), backgroundColor: stageColors(), borderColor: stageColors() }]));
    create("traffic", "trafficChart", makeChartConfig("doughnut", Object.keys(trafficMap), [{ label: "Origem", data: Object.values(trafficMap) }], { scales: {} }));
    create("owner", "ownerChart", makeChartConfig("bar", Object.keys(ownerMap), [{ label: "Receita fechada", data: Object.values(ownerMap) }], { indexAxis: "y" }));
    create("yearlyDaily", "yearlyDailyChart", makeChartConfig("line", yearLabels, [{
      label: "Contatos diários",
      data: yearDateKeys.map((key) => yearDayMap[key] || 0),
      tension: 0.2,
      fill: true,
      pointRadius: 0,
      pointHoverRadius: 4,
      borderWidth: 2
    }], {
      scales: {
        x: {
          ticks: {
            color: "#edf7f0",
            maxRotation: 0,
            autoSkip: false,
            callback: (value, index) => (monthStartIndexes.has(index) ? yearLabels[index] : "")
          },
          grid: { color: "rgba(255,255,255,0.08)" }
        },
        y: {
          ticks: { color: "#edf7f0", precision: 0 },
          grid: { color: "rgba(255,255,255,0.08)" },
          beginAtZero: true
        }
      }
    }));

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
    const delimiterScores = [
      { value: ";", score: (sample.match(/;/g) || []).length },
      { value: ",", score: (sample.match(/,/g) || []).length },
      { value: "\t", score: (sample.match(/\t/g) || []).length }
    ].sort((a, b) => b.score - a.score);
    const delimiter = delimiterScores[0]?.score ? delimiterScores[0].value : ";";

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

  async function readFileAsText(file) {
    const buffer = await file.arrayBuffer();
    const decoders = [
      new TextDecoder("utf-8", { fatal: false }),
      new TextDecoder("windows-1252", { fatal: false }),
      new TextDecoder("iso-8859-1", { fatal: false })
    ];

    for (const decoder of decoders) {
      const text = decoder.decode(buffer);
      if (/[A-Za-z0-9]/.test(text)) return text;
    }

    return new TextDecoder("utf-8", { fatal: false }).decode(buffer);
  }

  async function parseImportedRows(file) {
    const filename = String(file?.name || "").toLowerCase();

    if (filename.endsWith(".xlsx") || filename.endsWith(".xls")) {
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX?.read(buffer, { type: "array" });
      const firstSheetName = workbook?.SheetNames?.[0];
      const firstSheet = firstSheetName ? workbook.Sheets[firstSheetName] : null;
      const rows = firstSheet ? window.XLSX.utils.sheet_to_json(firstSheet, { defval: "" }) : [];
      const normalizedRows = rows.map((row) => {
        const next = {};
        Object.entries(row || {}).forEach(([key, value]) => {
          next[normalizeCsvHeader(key)] = value;
        });
        return next;
      });
      return { headers: Object.keys(normalizedRows[0] || {}), rows: normalizedRows };
    }

    const content = await readFileAsText(file);
    return parseCsv(content);
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
      stageIdByName(row.status) ||
      stageIdByName(row.etapa) ||
      stageIdByName(row.stage) ||
      state.stages[0]?.id ||
      null
    );
  }

  function parseMoney(value) {
    return parseMonetaryValue(value);
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

    downloadTextFile("leads.csv", `\uFEFF${csvRows.join("\n")}`, "text/csv;charset=utf-8");
  }

  async function importLeadsFromCsv(file) {
    if (!file) return;

    const { rows } = await parseImportedRows(file);

    if (!rows.length) {
      alert("O CSV está vazio ou inválido.");
      return;
    }

    const payload = rows
      .map((row) => {
        const name = row.nome || row.name || "";
        const contact = row.contato || row.contact || "";
        const owner = canAssignLeadOwner()
          ? (row.responsavel || row.vendedor || row.owner || getUserDisplayName())
          : getUserDisplayName();
        const startDate = row.data_inicio || row.start_date || row.data || "";
        const trafficType = row.origem || row.traffic_type || getLeadSourceNames()[0] || "Organico";
        const legacyText = String(row.observacoes || row.notes || "").trim();
        const importedValue = parseMoney(row.quantidade || row.valor || row.value || 0);
        const importedContractNumber = String(row.contrato || row.numero_contrato || row.contract_number || "").trim();
        const importedClosedAt = normalizeDateInput(row.data_fechamento || row.fechado_em || row.data || "") || "";
        const plan = String(row.plano || row.plan || "").trim();
        const planName = plan || (importedValue > 0 ? getDefaultPlanName(0) : "");
        const plans = planName
          ? [{
              name: planName,
              value: importedValue,
              contract_number: importedContractNumber,
              closed_at: importedClosedAt
            }]
          : [];

        return {
          stage_id: inferStageId(row),
          assigned_to: state.currentUser?.id || null,
          created_by: state.currentUser?.id || null,
          name: String(name).trim(),
          contact: String(contact).trim(),
          owner: String(owner).trim(),
          value: importedValue,
          start_date: normalizeDateInput(startDate) || startDate,
          social_source: String(row.rede_social || row.social_source || "").trim(),
          traffic_type: String(trafficType).trim(),
          notes: serializeLeadMeta({
            plan: planName,
            plans,
            legacyText,
            observations: legacyText ? [{ date: "", text: legacyText }] : []
          })
        };
      })
      .filter((lead) => lead.name && lead.contact && lead.stage_id && lead.traffic_type);

    if (!payload.length) {
      alert("Nenhuma linha valida encontrada. Use colunas como nome, contato, data_inicio, origem e pipeline.");
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
    applyRoleBasedUi();
    if (!canManageAdminAreas() && (state.activeView === "solicitacoes" || state.activeView === "configuracoes")) {
      bindView("funil");
    }
    populateFilters();
    renderStats();
    renderPipeline();
    renderLeadTable();
    renderTeam();
    renderStagesConfig();
    renderLeadSourcesConfig();
    renderRequests();
    renderHistoryText();
    if (state.activeView === "relatorios" && typeof window.Chart !== "undefined") {
      renderCharts();
    }
    bindGeneralActionEvents();
    requestAnimationFrame(updateStickyLayout);
  }

  function openLeadModal(lead = null) {
    closeAllModals();
    els.leadForm.reset();
    els.leadId.value = lead?.id || "";
    els.modalTitle.textContent = lead ? "Editar Lead" : "Novo Lead";
    els.ownerGroup?.classList.toggle("hidden", !canAssignLeadOwner());
    els.planGroup?.classList.toggle("hidden", !lead);

    els.name.value = lead?.name || "";
    els.contact.value = lead?.contact || "";
    els.owner.value = canAssignLeadOwner() ? (lead?.owner || getUserDisplayName()) : getUserDisplayName();
    if (els.value) els.value.value = "";
    els.startDate.value = lead?.start_date || "";
    els.stage.value = lead?.stage_id || state.stages[0]?.id || "";
    if (!state.stages.length) {
      alert("Nenhuma etapa encontrada no CRM. Crie uma etapa primeiro.");
      closeLeadModal();
      return;
    }
    els.socialSource.value = lead?.social_source || "";
    els.trafficType.value = lead?.traffic_type || getLeadSourceNames()[0] || "";
    state.modalPlans = getLeadPlans(lead).map((item) => ({ ...item }));
    if (lead && !state.modalPlans.length && Number(lead?.value || 0) > 0) {
      state.modalPlans = [{ name: getDefaultPlanName(0), value: Number(lead.value || 0) }];
    } else if (lead && !state.modalPlans.length) {
      state.modalPlans = [{ name: "Sem plano", value: 0 }];
    }
    state.modalObservations = getLeadObservations(lead);
    renderPlanItems();
    renderObservationItems();

    openModalOverlay(els.modalOverlay, "#name");
  }

  function closeLeadModal() {
    closeModalOverlay(els.modalOverlay);
  }

  function openStageModal(stage = null) {
    closeAllModals();
    els.stageForm.reset();
    els.stageId.value = stage?.id || "";
    els.stageModalTitle.textContent = stage ? "Editar pipeline" : "Adicionar pipeline";
    els.saveStageBtn.textContent = stage ? "Salvar alterações" : "Adicionar";
    els.stageName.value = stage?.name || "";
    els.stageColor.value = sanitizeHexColor(stage?.color);
    refreshStageTypeOptions(stage?.custom_stage_type ? `custom:${stage.custom_stage_type}` : (stage?.stage_type || "andamento"), stage?.custom_stage_type || "");
    updateStageColorPreview(els.stageColor.value);
    openModalOverlay(els.stageModalOverlay, "#stageName");
  }

  function closeStageModal() {
    closeModalOverlay(els.stageModalOverlay);
  }

  async function openHistoryModal() {
    if (!canViewHistory()) {
      alert("Somente administradores podem visualizar o historico completo.");
      return;
    }
    closeAllModals();
    els.historyText.textContent = state.historyLoaded ? els.historyText.textContent : "Carregando histórico...";
    renderHistoryText();
    openModalOverlay(els.historyModalOverlay);
    await loadHistory();
  }

  function closeHistoryModal() {
    closeModalOverlay(els.historyModalOverlay);
  }

  function requestAdminAuthorization(context) {
    state.permissionRequestContext = context || null;
    if (!state.permissionRequestContext) return;

    els.permissionModalTitle.textContent = state.permissionRequestContext.title || "Solicitar autorizacao";
    els.permissionModalText.textContent = state.permissionRequestContext.description || "Esta acao exige aprovacao de um administrador.";
    if (els.permissionRequestReason) els.permissionRequestReason.value = "";
    openModalOverlay(els.permissionModalOverlay, "#permissionRequestReason");
  }

  function closePermissionRequestModal() {
    state.permissionRequestContext = null;
    closeModalOverlay(els.permissionModalOverlay);
  }

  async function submitPermissionRequest() {
    if (!state.permissionRequestContext || !state.currentUser) return;

    const reason = String(els.permissionRequestReason?.value || "").trim();
    const payload = {
      request_type: state.permissionRequestContext.requestType || "manual_authorization",
      entity_type: state.permissionRequestContext.entityType || null,
      entity_id: state.permissionRequestContext.entityId ? String(state.permissionRequestContext.entityId) : null,
      title: state.permissionRequestContext.title || "Solicitacao operacional",
      description: state.permissionRequestContext.description || "",
      reason,
      status: ACCESS_STATUS.PENDING,
      payload: state.permissionRequestContext.payload || {},
      requested_by_id: state.currentUser.id,
      requested_by_name: getUserDisplayName(),
      requested_by_email: state.currentUser.email
    };

    const { error } = await state.supabase.from("admin_requests").insert([payload]);
    if (error) {
      alert(`Erro ao registrar solicitacao: ${error.message}`);
      return;
    }

    closePermissionRequestModal();
    alert("Solicitacao enviada para a aba Solicitacoes.");
    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  async function updateProfileRole(profileId, nextRole) {
    if (!canManageAdminAreas() || !profileId) return;

    const role = String(nextRole || USER_ROLE.USER).toLowerCase() === USER_ROLE.ADMIN
      ? USER_ROLE.ADMIN
      : USER_ROLE.USER;

    const { error } = await state.supabase
      .from("profiles")
      .update({ role })
      .eq("id", profileId);

    if (error) {
      alert(`Erro ao atualizar perfil: ${error.message}`);
      return;
    }

    if (state.profile?.id === profileId) {
      state.profile = { ...state.profile, role };
      applyRoleBasedUi();
    }

    await loadAppData({ includeProfiles: true });
  }

  async function deleteTeamMember(profileId) {
    if (!canManageAdminAreas() || !profileId) return;

    const profile = state.profiles.find((item) => item.id === profileId);
    if (!profile) {
      alert("Usuario nao encontrado.");
      return;
    }

    if (state.currentUser?.id === profileId) {
      alert("Voce nao pode excluir o proprio usuario.");
      return;
    }

    const label = profile.full_name || profile.email || "este usuario";
    if (!window.confirm(`Excluir ${label}? Esta acao remove o acesso ao CRM permanentemente.`)) {
      return;
    }

    const { data, error } = await state.supabase.functions.invoke("admin-delete-user", {
      body: { profileId }
    });

    const message = data?.error || error?.message;
    if (error || data?.error) {
      alert(`Erro ao excluir usuario: ${message || "Falha desconhecida."}`);
      return;
    }

    await logChange(
      "delete_user",
      "profile",
      profileId,
      `Usuario removido: ${profile.full_name || profile.email || profileId}`,
      { email: profile.email || null, role: profile.role || null }
    );

    await loadAppData({ includeProfiles: true });
    alert("Usuario excluido com sucesso.");
  }

  async function approveAccessRequest(requestId, role) {
    if (!canManageAdminAreas()) return;
    const request = state.accessRequests.find((item) => item.id === requestId);
    if (!request) return;

    const approvedRole = String(role || USER_ROLE.USER).toLowerCase() === USER_ROLE.ADMIN
      ? USER_ROLE.ADMIN
      : USER_ROLE.USER;

    const profileFilter = request.auth_user_id
      ? state.supabase.from("profiles").update({
          access_status: ACCESS_STATUS.APPROVED,
          role: approvedRole,
          approved_at: new Date().toISOString(),
          approved_by: state.currentUser.id
        }).eq("id", request.auth_user_id)
      : state.supabase.from("profiles").update({
          access_status: ACCESS_STATUS.APPROVED,
          role: approvedRole,
          approved_at: new Date().toISOString(),
          approved_by: state.currentUser.id
        }).eq("email", request.email);

    const { error: profileError } = await profileFilter;
    if (profileError) {
      alert(`Erro ao aprovar acesso: ${profileError.message}`);
      return;
    }

    const { error } = await state.supabase
      .from("access_requests")
      .update({
        status: ACCESS_STATUS.APPROVED,
        reviewed_at: new Date().toISOString(),
        reviewed_by: state.currentUser.id,
        approved_role: approvedRole
      })
      .eq("id", requestId);

    if (error) {
      alert(`Erro ao atualizar solicitacao: ${error.message}`);
      return;
    }

    await loadAppData({ includeProfiles: true });
  }

  async function rejectAccessRequest(requestId) {
    if (!canManageAdminAreas()) return;
    const request = state.accessRequests.find((item) => item.id === requestId);
    if (!request) return;

    const profileFilter = request.auth_user_id
      ? state.supabase.from("profiles").update({
          access_status: ACCESS_STATUS.REJECTED,
          approved_by: state.currentUser.id
        }).eq("id", request.auth_user_id)
      : state.supabase.from("profiles").update({
          access_status: ACCESS_STATUS.REJECTED,
          approved_by: state.currentUser.id
        }).eq("email", request.email);

    const { error: profileError } = await profileFilter;
    if (profileError) {
      alert(`Erro ao recusar acesso: ${profileError.message}`);
      return;
    }

    const { error } = await state.supabase
      .from("access_requests")
      .update({
        status: ACCESS_STATUS.REJECTED,
        reviewed_at: new Date().toISOString(),
        reviewed_by: state.currentUser.id
      })
      .eq("id", requestId);

    if (error) {
      alert(`Erro ao atualizar solicitacao: ${error.message}`);
      return;
    }

    await loadAppData({ includeProfiles: true });
  }

  async function resolveAdminRequest(requestId, action) {
    if (!canManageAdminAreas()) return;
    const request = state.adminRequests.find((item) => item.id === requestId);
    if (!request) return;

    if (action === "approve") {
      if (request.request_type === "delete_lead" && request.payload?.lead_id) {
        const { error: deleteError } = await state.supabase
          .from("leads")
          .delete()
          .eq("id", request.payload.lead_id);
        if (deleteError) {
          alert(`Erro ao excluir lead aprovado: ${deleteError.message}`);
          return;
        }
      }

      if (request.request_type === "bulk_delete_leads" && Array.isArray(request.payload?.lead_ids) && request.payload.lead_ids.length) {
        const { error: deleteError } = await state.supabase
          .from("leads")
          .delete()
          .in("id", request.payload.lead_ids);
        if (deleteError) {
          alert(`Erro ao excluir leads aprovados: ${deleteError.message}`);
          return;
        }
      }
    }

    const nextStatus = action === "approve" ? ACCESS_STATUS.APPROVED : ACCESS_STATUS.REJECTED;
    const { error } = await state.supabase
      .from("admin_requests")
      .update({
        status: nextStatus,
        reviewed_at: new Date().toISOString(),
        reviewed_by: state.currentUser.id
      })
      .eq("id", requestId);

    if (error) {
      alert(`Erro ao atualizar solicitacao operacional: ${error.message}`);
      return;
    }

    await loadAppData({ includeProfiles: true });
  }

  async function addLeadSource() {
    if (!canManageLeadSources()) return;

    const name = String(els.leadSourceName?.value || "").trim();
    if (!name) {
      alert("Informe o nome da origem do lead.");
      return;
    }

    const { error } = await state.supabase
      .from("lead_source_catalog")
      .upsert({ name, created_by: state.currentUser?.id || null }, { onConflict: "name" });

    if (error && !isMissingRelationError(error)) {
      alert(`Erro ao salvar origem: ${error.message}`);
      return;
    }

    if (els.leadSourceName) els.leadSourceName.value = "";
    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  async function editLeadSource(sourceName) {
    if (!canManageLeadSources()) return;

    const currentName = String(sourceName || "").trim();
    if (!currentName) return;
    const nextName = window.prompt("Novo nome para a origem do lead:", currentName);
    if (!nextName || nextName.trim() === currentName) return;

    const { error: leadError } = await state.supabase
      .from("leads")
      .update({ traffic_type: nextName.trim() })
      .eq("traffic_type", currentName);

    if (leadError) {
      alert(`Erro ao atualizar leads: ${leadError.message}`);
      return;
    }

    const { error: updateError } = await state.supabase
      .from("lead_source_catalog")
      .update({ name: nextName.trim() })
      .eq("name", currentName);

    if (updateError && !isMissingRelationError(updateError)) {
      alert(`Erro ao atualizar origem: ${updateError.message}`);
      return;
    }

    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  async function deleteLeadSource(sourceName) {
    if (!canManageLeadSources()) return;

    const currentName = String(sourceName || "").trim();
    if (!currentName) return;

    const fallbackName = getLeadSourceNames().find((item) => item !== currentName) || DEFAULT_LEAD_SOURCES[0];
    if (!fallbackName) {
      alert("Mantenha ao menos uma origem do lead disponivel.");
      return;
    }

    if (!confirm(`Excluir a origem "${currentName}"? Leads existentes serao migrados para "${fallbackName}".`)) return;

    const { error: leadError } = await state.supabase
      .from("leads")
      .update({ traffic_type: fallbackName })
      .eq("traffic_type", currentName);

    if (leadError) {
      alert(`Erro ao atualizar leads vinculados: ${leadError.message}`);
      return;
    }

    const { error } = await state.supabase
      .from("lead_source_catalog")
      .delete()
      .eq("name", currentName);

    if (error && !isMissingRelationError(error)) {
      alert(`Erro ao excluir origem: ${error.message}`);
      return;
    }

    await loadAppData({ includeProfiles: state.profilesLoaded });
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
    const normalizedModalPlans = state.modalPlans.map((item) => {
      const name = String(item?.name || "").trim();
      const rawValue = String(item?.value ?? "").trim();
      const supportsClosingDetails = !isNoPlanName(name) && Number(parseMonetaryValue(item?.value)) > 0;
      return {
        name,
        value: isNoPlanName(name) && rawValue === "" ? 0 : item?.value,
        contract_number: supportsClosingDetails ? String(item?.contract_number || "").trim() : "",
        closed_at: supportsClosingDetails ? (normalizeDateInput(item?.closed_at || "") || "") : ""
      };
    });
    const draftPlans = existingLead
      ? cleanPlanList(normalizedModalPlans.map((item) => ({
          name: item.name,
          value: item.value,
          contract_number: item.contract_number,
          closed_at: item.closed_at
        })))
      : [];
    const leadValue = getPlansTotalValue(draftPlans);
    const draftObservations = cleanObservationList(state.modalObservations);
    const resolvedOwner = canAssignLeadOwner()
      ? els.owner.value.trim()
      : (existingLead?.owner || getUserDisplayName());

    const invalidPlan = normalizedModalPlans.find((item) => item.name && !isNoPlanName(item.name) && String(item?.value ?? "").trim() === "");
    if (existingLead && invalidPlan) return alert("Ao adicionar um plano, informe tambem o valor.");

    const payload = {
      assigned_to: state.currentUser.id,
      created_by: existingLead?.created_by || state.currentUser.id,
      stage_id: els.stage.value,
      name: els.name.value.trim(),
      contact: els.contact.value.trim(),
      owner: resolvedOwner,
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
    if (!payload.name || !payload.contact || !payload.owner || !payload.start_date || !payload.traffic_type) {
      return alert("Preencha os campos obrigatorios.");
    }

    let error;
    let savedLeadId = els.leadId.value || null;

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
      savedLeadId = data?.id || null;

      await logChange(
        "insert",
        "lead",
        data?.id,
        `Lead "${payload.name}" foi criado por ${getUserDisplayName()}.`,
        payload
      );
    }

    try {
      await syncPlanValuesAcrossLeads(draftPlans, savedLeadId);
    } catch (syncError) {
      alert(`Lead salvo, mas não foi possível sincronizar o valor do plano nos outros leads: ${syncError.message}`);
    }

    closeLeadModal();
    await loadAppData({ includeProfiles: state.profilesLoaded });
  }

  async function submitStage(event) {
    event.preventDefault();

    if (!canManageStages()) {
      alert("Somente administradores podem alterar pipelines.");
      return;
    }

    const selectedType = String(els.stageType.value || "andamento").trim();
    const customStageTypeInput = String(els.customStageType?.value || "").trim();
    const existingCustomType = selectedType.startsWith("custom:") ? selectedType.replace(/^custom:/, "") : "";
    const isCreatingNewCustom = selectedType === "personalizado";
    const isEditingExistingCustom = !!existingCustomType;
    const desiredCustomType = isCreatingNewCustom
      ? customStageTypeInput
      : (isEditingExistingCustom ? (customStageTypeInput || existingCustomType) : "");
    const payload = {
      name: els.stageName.value.trim(),
      color: sanitizeHexColor(els.stageColor.value),
      stage_type: (isCreatingNewCustom || isEditingExistingCustom)
        ? "personalizado"
        : (["andamento", "fechado", "cancelado", "espera"].includes(selectedType) ? selectedType : "andamento"),
      custom_stage_type: (isCreatingNewCustom || isEditingExistingCustom) ? desiredCustomType : null,
      position: els.stageId.value ? (state.stages.find((s) => s.id === els.stageId.value)?.position ?? state.stages.length) : state.stages.length,
      created_by: state.currentUser.id
    };

    if (!payload.name) return alert("Informe o nome da etapa.");
    if (payload.stage_type === "personalizado" && !payload.custom_stage_type) return alert("Informe o tipo personalizado.");

    if (false && payload.stage_type === "personalizado") {
      const savedType = await saveCustomStageType(payload.custom_stage_type);
      if (!savedType) return alert("Não foi possível salvar o novo tipo personalizado.");
    }

    const customTypes = getCustomStageTypes();
    const duplicateCustom = customTypes.find((item) =>
      normalizeStageTypeNameKey(item) === normalizeStageTypeNameKey(payload.custom_stage_type) &&
      normalizeStageTypeNameKey(item) !== normalizeStageTypeNameKey(existingCustomType)
    );

    if (isCreatingNewCustom && duplicateCustom) {
      return alert("Esse tipo personalizado ja existe. Selecione-o na lista para editar.");
    }

    if (isEditingExistingCustom && desiredCustomType !== existingCustomType) {
      if (duplicateCustom) return alert("Ja existe outro tipo com esse nome. Selecione-o na lista para editar.");
      const renamed = await renameCustomStageType(existingCustomType, desiredCustomType);
      if (!renamed) return alert("Nao foi possivel atualizar o tipo personalizado.");
    } else if (isCreatingNewCustom) {
      const savedType = await saveCustomStageType(payload.custom_stage_type);
      if (!savedType) return alert("Nao foi possivel salvar o novo tipo personalizado.");
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
    if (!canDeleteLeads()) {
      requestAdminAuthorization({
        requestType: "delete_lead",
        title: "Solicitar exclusao de lead",
        description: `Voce nao tem permissao para excluir o lead "${lead.name}". Sua solicitacao sera enviada para o administrador.`,
        entityType: "lead",
        entityId: lead.id,
        payload: {
          lead_id: lead.id,
          lead_name: lead.name,
          lead_owner: lead.owner || null
        }
      });
      return;
    }
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
    if (!canManageStages()) {
      alert("Somente administradores podem editar pipelines.");
      return;
    }
    const stage = state.stages.find((x) => x.id === id);
    if (!stage) return;
    openStageModal(stage);
  }

  async function deleteStage(id) {
    const stage = state.stages.find((x) => x.id === id);
    const fallback = state.stages.find((x) => x.id !== id);

    if (!stage) return;
    if (!canManageStages()) {
      requestAdminAuthorization({
        requestType: "delete_stage",
        title: "Solicitar exclusao de pipeline",
        description: `Voce nao tem permissao para excluir a etapa "${stage.name}". Sua solicitacao sera enviada para o administrador.`,
        entityType: "stage",
        entityId: stage.id,
        payload: {
          stage_id: stage.id,
          stage_name: stage.name
        }
      });
      return;
    }
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

    const redirectTo = getAuthRedirectUrl("#type=recovery");

    if (!redirectTo) {
      return setMessage(
        els.authMessage,
        "Abra o CRM por uma URL http/https para usar e-mails de autenticacao. Arquivo local nao funciona com redirect de autenticacao.",
        true
      );
    }

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
    if ((name === "solicitacoes" || name === "configuracoes") && !canManageAdminAreas()) {
      name = "funil";
    }

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
      relatorios: ["Relatorios", "Acompanhe os resultados do CRM compartilhado."],
      equipe: ["Equipe", "Usuarios que ja interagiram no CRM."],
      solicitacoes: ["Solicitacoes", "Aprove acessos e gerencie pedidos administrativos."],
      configuracoes: ["Configuracoes", "Gerencie estruturas compartilhadas e listas dinamicas do CRM."]
    };

    els.pageTitle.textContent = titles[name][0];
    els.pageSubtitle.textContent = titles[name][1];
    els.sidebar.classList.remove("open");
    setMobileFiltersOpen(false);
    requestAnimationFrame(updateStickyLayout);

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
    normalizeMobileFilterTexts();

    document.querySelectorAll(".tab-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        if (btn.classList.contains("hidden")) return;
        const targetPanel = $(`${btn.dataset.tab}Form`);
        if (!targetPanel || targetPanel.classList.contains("hidden")) return;
        document.querySelectorAll(".tab-btn").forEach((b) => b.classList.toggle("active", b === btn));
        document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.remove("active"));
        targetPanel.classList.add("active");
      });
    });

    els.loginForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      setMessage(els.authMessage, "");

      const { data, error } = await state.supabase.auth.signInWithPassword({
        email: $("loginEmail").value.trim(),
        password: $("loginPassword").value.trim()
      });

      if (error) return setMessage(els.authMessage, getAuthErrorMessage(error, "Nao foi possivel fazer login."), true);

      state.currentUser = data?.user || null;
      if (!state.currentUser) {
        return setMessage(els.authMessage, "Nao foi possivel iniciar a sessao. Tente novamente.", true);
      }

      await ensureProfile();
      if (!(await enforceApprovedSession())) return;
      await enterApp();
    });

    els.registerForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      setMessage(els.authMessage, "");

      if (!state.security.allowSelfRegistration) {
        return setMessage(els.authMessage, getSignupRestrictionMessage(), true);
      }

      const email = $("registerEmail").value.trim();
      if (!isSignupEmailAllowed(email)) {
        return setMessage(els.authMessage, getSignupRestrictionMessage(), true);
      }
      const password = $("registerPassword").value.trim();
      const full_name = $("registerName").value.trim();
      const emailRedirectTo = getAuthRedirectUrl();

      const { data, error } = await state.supabase.auth.signUp({
        email,
        password,
        options: {
          data: { full_name },
          ...(emailRedirectTo ? { emailRedirectTo } : {})
        }
      });

      if (error) return setMessage(els.authMessage, getAuthErrorMessage(error), true);

      const createdUser = data.user || data.session?.user || null;
      const hasSession = Boolean(data.session?.access_token);
      state.currentUser = hasSession ? createdUser : null;

      if (createdUser) {
        if (hasSession) {
          await ensureProfile();
        }

        const requestPayload = {
          full_name,
          email,
          status: ACCESS_STATUS.PENDING
        };

        if (hasSession) {
          requestPayload.auth_user_id = createdUser.id;
        }

        const { error: requestError } = await state.supabase
          .from("access_requests")
          .insert([requestPayload]);

        if (requestError && !isMissingRelationError(requestError)) {
          if (isDuplicateKeyError(requestError)) {
            // The backend may create the access request automatically on signup.
          } else {
            return setMessage(els.authMessage, `Nao foi possivel registrar a solicitacao: ${requestError.message}`, true);
          }
        }

        try {
          await state.supabase.functions.invoke("notify-admin-access-request", {
            body: { email, full_name }
          });
        } catch (_error) {
          // Optional function: the request remains recorded even when notification is unavailable.
        }

        if (hasSession) {
          await state.supabase.auth.signOut();
        }
      }

      if (!createdUser) {
        return setMessage(els.authMessage, "O cadastro foi recebido, mas nao foi possivel confirmar a solicitacao de acesso. Tente novamente.", true);
      }

      state.currentUser = null;
      resetAppState();
      setMessage(els.authMessage, "Solicitacao enviada. Aguarde a aprovacao do administrador para acessar o CRM.");
      document.querySelector('[data-tab="login"]').click();
      $("loginEmail").value = email;
      $("loginPassword").value = "";
      $("registerForm").reset();
    });

    els.forgotPasswordBtn.addEventListener("click", sendResetPasswordEmail);
    els.updatePasswordBtn.addEventListener("click", updateRecoveredPassword);

    els.logoutBtn.addEventListener("click", async () => {
      await state.supabase.auth.signOut();
      resetAppState();
      showScreen("authScreen");
    });

    els.mobileMenuBtn?.addEventListener("click", () => {
      setMobileFiltersOpen(false);
      els.sidebar.classList.toggle("open");
    });

    document.querySelectorAll(".menu-item").forEach((btn) => {
      btn.addEventListener("click", () => bindView(btn.dataset.view));
    });

    els.searchInput.addEventListener("input", () => {
      renderAll();
    });

    els.mobileFiltersBtn?.addEventListener("click", () => {
      els.sidebar?.classList.remove("open");
      setMobileFiltersOpen(els.mobileFiltersPanel?.classList.contains("hidden"));
    });

    els.mobileOwnerFilter?.addEventListener("change", () => {
      if (els.ownerFilter) els.ownerFilter.value = els.mobileOwnerFilter.value;
      renderAll();
    });

    els.mobileMonthFilter?.addEventListener("change", () => {
      if (els.monthFilter) els.monthFilter.value = els.mobileMonthFilter.value;
      renderAll();
    });

    els.mobileLeadSourceFilter?.addEventListener("change", () => {
      if (els.leadSourceFilter) els.leadSourceFilter.value = els.mobileLeadSourceFilter.value;
      renderAll();
    });

    els.mobileClearFiltersBtn?.addEventListener("click", () => {
      if (els.ownerFilter) els.ownerFilter.value = "";
      if (els.monthFilter) els.monthFilter.value = "";
      if (els.leadSourceFilter) els.leadSourceFilter.value = "";
      syncMobileFilterControls();
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

    els.leadSourceFilterBtn?.addEventListener("click", (e) => {
      e.stopPropagation();
      const isOpen = els.leadSourceFilterDropdown?.classList.contains("open");
      closeFilterDropdowns(els.leadSourceFilterDropdown);
      setFilterDropdownOpen(els.leadSourceFilterDropdown, els.leadSourceFilterBtn, els.leadSourceFilterMenu, !isOpen);
    });

    document.addEventListener("click", (e) => {
      if (els.ownerFilterDropdown?.contains(e.target) || els.monthFilterDropdown?.contains(e.target) || els.leadSourceFilterDropdown?.contains(e.target)) return;
      if (els.mobileFiltersBtn?.contains(e.target) || els.mobileFiltersPanel?.contains(e.target)) return;
      closeFilterDropdowns();
      if (isCompactViewport()) setMobileFiltersOpen(false);
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
    els.addStageBtn?.addEventListener("click", openStageModal);
    els.historyBtn?.addEventListener("click", openHistoryModal);
    els.addLeadSourceBtn?.addEventListener("click", addLeadSource);

    els.closeModalBtn.addEventListener("click", closeLeadModal);
    els.cancelBtn.addEventListener("click", closeLeadModal);
    els.leadForm.addEventListener("submit", submitLead);

    els.closeStageModalBtn.addEventListener("click", closeStageModal);
    els.cancelStageBtn.addEventListener("click", closeStageModal);
    els.stageForm.addEventListener("submit", submitStage);
    els.stageColor.addEventListener("input", (e) => updateStageColorPreview(e.target.value));
    els.stageType.addEventListener("change", toggleCustomStageTypeField);
    els.removeCustomTypeBtn.addEventListener("click", removeCurrentCustomStageType);
    els.removeSelectedStageTypeBtn?.addEventListener("click", removeCurrentCustomStageType);
    els.closePermissionModalBtn?.addEventListener("click", closePermissionRequestModal);
    els.cancelPermissionRequestBtn?.addEventListener("click", closePermissionRequestModal);
    els.submitPermissionRequestBtn?.addEventListener("click", submitPermissionRequest);
    els.savedStageTypes?.addEventListener("click", async (event) => {
      const selectBtn = event.target.closest("[data-stage-type-value]");
      if (selectBtn) {
        const type = String(selectBtn.dataset.stageTypeValue || "").trim();
        if (!type) return;
        els.stageType.value = type;
        toggleCustomStageTypeField();
      }
    });
    setupPlanListEvents();
    setupObservationListEvents();
    attachPipelineScrollEvents();
    window.addEventListener("resize", () => {
      if (!isCompactViewport()) setMobileFiltersOpen(false);
      requestAnimationFrame(() => {
        updateStickyLayout();
        syncPipelineScrollBars();
      });
    });

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

    document.addEventListener("click", async (event) => {
      const accessActionBtn = event.target.closest("[data-access-action]");
      if (accessActionBtn) {
        const requestId = accessActionBtn.dataset.id;
        const roleSelect = document.querySelector(`[data-access-role="${requestId}"]`);
        if (accessActionBtn.dataset.accessAction === "approve") {
          await approveAccessRequest(requestId, roleSelect?.value || USER_ROLE.USER);
        }
        if (accessActionBtn.dataset.accessAction === "reject") {
          await rejectAccessRequest(requestId);
        }
        return;
      }

      const adminRequestBtn = event.target.closest("[data-admin-request-action]");
      if (adminRequestBtn) {
        await resolveAdminRequest(adminRequestBtn.dataset.id, adminRequestBtn.dataset.adminRequestAction);
        return;
      }

      const sourceBtn = event.target.closest("[data-source-action]");
      if (sourceBtn) {
        const sourceName = sourceBtn.dataset.sourceName;
        if (sourceBtn.dataset.sourceAction === "edit") await editLeadSource(sourceName);
        if (sourceBtn.dataset.sourceAction === "delete") await deleteLeadSource(sourceName);
        return;
      }

      const teamActionBtn = event.target.closest("[data-team-action]");
      if (teamActionBtn) {
        if (teamActionBtn.dataset.teamAction === "delete") {
          await deleteTeamMember(teamActionBtn.dataset.profileId);
        }
      }
    });

    document.addEventListener("change", async (event) => {
      const roleSelect = event.target.closest(".team-role-select");
      if (roleSelect) {
        await updateProfileRole(roleSelect.dataset.profileId, roleSelect.value);
      }
    });

    els.closeHistoryModalBtn.addEventListener("click", closeHistoryModal);
    els.refreshHistoryBtn.addEventListener("click", async () => {
      els.historyText.textContent = "Carregando histórico...";
      await loadHistory(true);
    });

    bindOverlayDismiss(els.modalOverlay, closeLeadModal);
    bindOverlayDismiss(els.stageModalOverlay, closeStageModal);
    bindOverlayDismiss(els.historyModalOverlay, closeHistoryModal);
    bindOverlayDismiss(els.permissionModalOverlay, closePermissionRequestModal);

    state.supabase.auth.onAuthStateChange((event, session) => {
      window.setTimeout(() => {
        handleAuthStateChange(event, session).catch((error) => {
          console.error("Erro ao processar evento de autenticação:", error);
        });
      }, 0);
    });
  }

  async function init() {
    try {
      runPeriodicStorageCleanup();
      state.security = getSecurityConfig();
      applySecurityConfigToUi();
      createClient();
      bindEvents();
      requestAnimationFrame(updateStickyLayout);
      await bootstrap();
    } catch (error) {
      console.error(error);
      els.bootScreen.textContent = "Erro ao iniciar o sistema. Verifique o config.js e o Supabase.";
    }
  }

  init();
})();
