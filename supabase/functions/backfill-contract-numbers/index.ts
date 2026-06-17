import { createClient } from "supabase";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type"
};

type ContractItem = {
  contact: string;
  contract_number: string;
  name?: string;
};

type LeadMeta = {
  plan: string;
  plans: Array<{ name: string; value: number; contract_number: string; closed_at: string }>;
  legacyText: string;
  observations: Array<{ date: string; text: string }>;
  contract_number: string;
  referral_name: string;
  referral_sector: string;
};

function getEnv(name: string) {
  const value = Deno.env.get(name);
  if (!value) throw new Error(`Missing environment variable: ${name}`);
  return value;
}

function normalizeWhitespace(value: unknown) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function normalizeContact(value: unknown) {
  const digits = String(value ?? "").replace(/\D/g, "");
  if (!digits) return "";
  return digits.startsWith("55") && digits.length > 11 ? digits.slice(2) : digits;
}

function normalizeDateInput(value: unknown) {
  const raw = normalizeWhitespace(value);
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  return "";
}

function cleanPlanList(plans: unknown): LeadMeta["plans"] {
  return (Array.isArray(plans) ? plans : [])
    .map((item) => {
      const record = item as Record<string, unknown>;
      const name = normalizeWhitespace(record?.name);
      const rawValue = Number(record?.value ?? 0);
      return {
        name,
        value: Number.isFinite(rawValue) ? rawValue : 0,
        contract_number: normalizeWhitespace(record?.contract_number ?? record?.contractNumber),
        closed_at: normalizeDateInput(record?.closed_at ?? record?.closedAt)
      };
    })
    .filter((item) => item.name);
}

function cleanObservationList(observations: unknown): LeadMeta["observations"] {
  return (Array.isArray(observations) ? observations : [])
    .map((item) => {
      const record = item as Record<string, unknown>;
      return {
        date: normalizeDateInput(record?.date),
        text: normalizeWhitespace(record?.text)
      };
    })
    .filter((item) => item.text);
}

function getDefaultPlanName(index = 0) {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  const safeIndex = Math.max(0, Number(index) || 0);
  const letter = alphabet[safeIndex % alphabet.length];
  const suffix = Math.floor(safeIndex / alphabet.length);
  return `Plano ${letter}${suffix ? suffix + 1 : ""}`;
}

function getLeadMeta(rawNotes: unknown, leadValue = 0): LeadMeta {
  const raw = String(rawNotes ?? "");
  const prefix = "__CRM_META__";

  if (!raw.startsWith(prefix)) {
    const text = raw.trim();
    return {
      plan: "",
      plans: [],
      legacyText: text,
      observations: text ? [{ date: "", text }] : [],
      contract_number: "",
      referral_name: "",
      referral_sector: ""
    };
  }

  try {
    const parsed = JSON.parse(raw.slice(prefix.length)) as Record<string, unknown>;
    const legacyText = normalizeWhitespace(parsed?.legacyText);
    const observations = cleanObservationList(parsed?.observations);
    const plans = cleanPlanList(parsed?.plans);
    const legacyPlan = normalizeWhitespace(parsed?.plan);
    const contractNumber = normalizeWhitespace(parsed?.contract_number ?? parsed?.contractNumber);
    const referralName = normalizeWhitespace(parsed?.referral_name);
    const referralSector = normalizeWhitespace(parsed?.referral_sector);

    if (!plans.length && legacyPlan) {
      plans.push({ name: legacyPlan, value: Number(leadValue || 0) || 0, contract_number: "", closed_at: "" });
    } else if (!plans.length && Number(leadValue || 0) > 0) {
      plans.push({ name: getDefaultPlanName(0), value: Number(leadValue || 0), contract_number: "", closed_at: "" });
    }
    if (!observations.length && legacyText) observations.push({ date: "", text: legacyText });

    return {
      plan: legacyPlan || plans[0]?.name || "",
      plans,
      legacyText,
      observations,
      contract_number: contractNumber,
      referral_name: referralName,
      referral_sector: referralSector
    };
  } catch {
    const text = raw.trim();
    return {
      plan: "",
      plans: [],
      legacyText: text,
      observations: text ? [{ date: "", text }] : [],
      contract_number: "",
      referral_name: "",
      referral_sector: ""
    };
  }
}

function serializeLeadMeta(meta: Partial<LeadMeta> = {}) {
  const legacyText = normalizeWhitespace(meta.legacyText);
  const plans = cleanPlanList(meta.plans);
  const observations = cleanObservationList(meta.observations);
  const plan = plans[0]?.name || normalizeWhitespace(meta.plan);
  const contractNumber = normalizeWhitespace(meta.contract_number);
  const referralName = normalizeWhitespace(meta.referral_name);
  const referralSector = normalizeWhitespace(meta.referral_sector);

  if (!plan && !plans.length && !observations.length && !contractNumber && !referralName && !referralSector) {
    return legacyText;
  }

  return `__CRM_META__${JSON.stringify({
    plan,
    plans,
    legacyText,
    observations,
    contract_number: contractNumber,
    referral_name: referralName,
    referral_sector: referralSector
  })}`;
}

Deno.serve(async (request) => {
  if (request.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  if (request.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed." }), {
      status: 405,
      headers: { ...corsHeaders, "Content-Type": "application/json" }
    });
  }

  try {
    const expectedAuthorization = `Bearer ${getEnv("SUPABASE_ANON_KEY")}`;
    const requestAuthorization = request.headers.get("Authorization");

    if (!requestAuthorization || requestAuthorization !== expectedAuthorization) {
      return new Response(JSON.stringify({ error: "Forbidden." }), {
        status: 403,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    const body = await request.json().catch(() => ({}));
    const dryRun = Boolean(body?.dry_run);
    const rawItems = Array.isArray(body?.items) ? body.items : [];
    const itemMap = new Map<string, ContractItem>();

    for (const entry of rawItems) {
      const contact = normalizeContact((entry as Record<string, unknown>)?.contact);
      const contractNumber = normalizeWhitespace((entry as Record<string, unknown>)?.contract_number);
      const name = normalizeWhitespace((entry as Record<string, unknown>)?.name);
      if (!contact || !contractNumber) continue;
      if (!itemMap.has(contact)) itemMap.set(contact, { contact, contract_number: contractNumber, name });
    }

    if (!itemMap.size) {
      return new Response(JSON.stringify({ error: "No valid contract items received." }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    const adminClient = createClient(getEnv("SUPABASE_URL"), getEnv("SUPABASE_SERVICE_ROLE_KEY"));
    const contacts = [...itemMap.keys()];
    const leads: Array<Record<string, unknown>> = [];

    for (let index = 0; index < contacts.length; index += 100) {
      const chunk = contacts.slice(index, index + 100);
      const { data, error } = await adminClient
        .from("leads")
        .select("id, name, contact, value, notes")
        .in("contact", chunk);

      if (error) throw error;
      leads.push(...(data || []));
    }

    const updated: Array<Record<string, unknown>> = [];
    const alreadyFilled: Array<Record<string, unknown>> = [];
    const conflicts: Array<Record<string, unknown>> = [];
    const notFound = contacts.filter((contact) => !leads.some((lead) => normalizeContact(lead.contact) === contact));

    for (const lead of leads) {
      const contact = normalizeContact(lead.contact);
      const match = itemMap.get(contact);
      if (!match) continue;

      const meta = getLeadMeta(lead.notes, Number(lead.value || 0));
      const existingContracts = cleanPlanList(meta.plans).map((item) => item.contract_number).filter(Boolean);
      const existingContract = existingContracts[0] || normalizeWhitespace(meta.contract_number);

      if (existingContract) {
        const bucket = existingContract === match.contract_number ? alreadyFilled : conflicts;
        bucket.push({
          id: lead.id,
          name: lead.name,
          contact,
          existing_contract_number: existingContract,
          incoming_contract_number: match.contract_number
        });
        continue;
      }

      meta.contract_number = match.contract_number;
      const nextPlans = cleanPlanList(meta.plans);
      const planIndex = nextPlans.findIndex((item) => normalizeWhitespace(item.name).toLowerCase() !== "sem plano");
      if (planIndex >= 0 && !nextPlans[planIndex].contract_number) {
        nextPlans[planIndex].contract_number = match.contract_number;
      }

      const nextNotes = serializeLeadMeta({
        ...meta,
        plans: nextPlans
      });

      if (!dryRun) {
        const { error } = await adminClient
          .from("leads")
          .update({ notes: nextNotes })
          .eq("id", String(lead.id));

        if (error) throw error;
      }

      updated.push({
        id: lead.id,
        name: lead.name,
        contact,
        contract_number: match.contract_number
      });
    }

    return new Response(JSON.stringify({
      ok: true,
      dry_run: dryRun,
      requested: itemMap.size,
      matched: leads.length,
      updated: updated.length,
      already_filled: alreadyFilled.length,
      conflicts: conflicts.length,
      not_found: notFound.length,
      samples: {
        updated: updated.slice(0, 10),
        already_filled: alreadyFilled.slice(0, 10),
        conflicts: conflicts.slice(0, 10),
        not_found: notFound.slice(0, 10)
      }
    }), {
      status: 200,
      headers: { ...corsHeaders, "Content-Type": "application/json" }
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: error instanceof Error ? error.message : "Unknown error" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" }
    });
  }
});
