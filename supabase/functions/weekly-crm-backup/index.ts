import { createClient } from "supabase";
import * as XLSX from "xlsx";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type"
};

const BUCKET_NAME = "crm-backups";
const CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const TABLES = [
  { name: "stages", sheet: "pipelines", orderBy: "position", ascending: true },
  { name: "leads", sheet: "leads", orderBy: "created_at", ascending: false },
  { name: "profiles", sheet: "usuarios", orderBy: "full_name", ascending: true },
  { name: "stage_type_catalog", sheet: "tipos_pipeline", orderBy: "name", ascending: true },
  { name: "lead_source_catalog", sheet: "origens_lead", orderBy: "name", ascending: true },
  { name: "access_requests", sheet: "solicitacoes_acesso", orderBy: "created_at", ascending: false },
  { name: "admin_requests", sheet: "solicitacoes_admin", orderBy: "created_at", ascending: false },
  { name: "change_history", sheet: "historico", orderBy: "created_at", ascending: false }
];

function getEnv(name: string) {
  const value = Deno.env.get(name);
  if (!value) throw new Error(`Missing environment variable: ${name}`);
  return value;
}

function sanitizeCellValue(value: unknown): string | number | boolean | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  return JSON.stringify(value);
}

function sanitizeRows(rows: Record<string, unknown>[]) {
  return rows.map((row) => {
    const next: Record<string, string | number | boolean | null> = {};
    Object.entries(row || {}).forEach(([key, value]) => {
      next[key] = sanitizeCellValue(value);
    });
    return next;
  });
}

function buildWorkbook(dataByTable: Record<string, Record<string, unknown>[]>, rowCounts: Record<string, number>) {
  const workbook = XLSX.utils.book_new();

  const metaRows = [
    { campo: "gerado_em_utc", valor: new Date().toISOString() },
    { campo: "bucket", valor: BUCKET_NAME },
    { campo: "total_abas", valor: TABLES.length }
  ];

  Object.entries(rowCounts).forEach(([table, count]) => {
    metaRows.push({ campo: `${table}_linhas`, valor: count });
  });

  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(metaRows), "meta");

  TABLES.forEach(({ name, sheet }) => {
    const rows = sanitizeRows(dataByTable[name] || []);
    const safeRows = rows.length ? rows : [{ sem_dados: "" }];
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(safeRows), sheet);
  });

  return XLSX.write(workbook, { bookType: "xlsx", type: "array" });
}

async function ensureBucket(adminClient: ReturnType<typeof createClient>) {
  const { data: buckets, error: listError } = await adminClient.storage.listBuckets();
  if (listError) throw listError;

  const exists = (buckets || []).some((bucket) => bucket.name === BUCKET_NAME);
  if (exists) return;

  const { error: createError } = await adminClient.storage.createBucket(BUCKET_NAME, {
    public: false,
    allowedMimeTypes: [CONTENT_TYPE],
    fileSizeLimit: "20MB"
  });

  if (createError && !String(createError.message || "").toLowerCase().includes("already")) {
    throw createError;
  }
}

async function insertRun(adminClient: ReturnType<typeof createClient>, triggerSource: string) {
  const { data, error } = await adminClient
    .from("crm_backup_runs")
    .insert([{ trigger_source: triggerSource, status: "running" }])
    .select("id")
    .single();

  if (error) throw error;
  return data.id as string;
}

async function updateRun(
  adminClient: ReturnType<typeof createClient>,
  runId: string,
  payload: Record<string, unknown>
) {
  const { error } = await adminClient
    .from("crm_backup_runs")
    .update({ ...payload, finished_at: new Date().toISOString() })
    .eq("id", runId);

  if (error) throw error;
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

    const triggerSource = String((await request.json().catch(() => ({})))?.trigger || "manual_function");
    const adminClient = createClient(getEnv("SUPABASE_URL"), getEnv("SUPABASE_SERVICE_ROLE_KEY"));
    const runId = await insertRun(adminClient, triggerSource);

    try {
      await ensureBucket(adminClient);

      const dataByTable: Record<string, Record<string, unknown>[]> = {};
      const rowCounts: Record<string, number> = {};

      for (const table of TABLES) {
        let query = adminClient.from(table.name).select("*");
        if (table.orderBy) {
          query = query.order(table.orderBy, { ascending: table.ascending });
        }

        const { data, error } = await query;
        if (error) throw new Error(`Failed to load ${table.name}: ${error.message}`);

        dataByTable[table.name] = (data || []) as Record<string, unknown>[];
        rowCounts[table.name] = dataByTable[table.name].length;
      }

      const workbookBuffer = buildWorkbook(dataByTable, rowCounts);
      const generatedAt = new Date().toISOString().replace(/[:.]/g, "-");
      const path = `weekly/crm-backup-${generatedAt}.xlsx`;
      const fileBytes = new Uint8Array(workbookBuffer);

      const { data: uploadData, error: uploadError } = await adminClient.storage
        .from(BUCKET_NAME)
        .upload(path, fileBytes, {
          contentType: CONTENT_TYPE,
          upsert: false
        });

      if (uploadError) throw uploadError;

      await updateRun(adminClient, runId, {
        status: "success",
        storage_path: uploadData.path,
        file_size_bytes: fileBytes.byteLength,
        row_counts: rowCounts,
        error_message: null
      });

      return new Response(JSON.stringify({
        ok: true,
        path: uploadData.path,
        rowCounts
      }), {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    } catch (error) {
      await updateRun(adminClient, runId, {
        status: "error",
        row_counts: {},
        error_message: error instanceof Error ? error.message : "Unknown error"
      });
      throw error;
    }
  } catch (error) {
    return new Response(JSON.stringify({ error: error instanceof Error ? error.message : "Unknown error" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" }
    });
  }
});
