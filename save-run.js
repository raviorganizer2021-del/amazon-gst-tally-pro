module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  const supabaseUrl = process.env.SUPABASE_URL;
  const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  const tableName = process.env.SUPABASE_TABLE || "gst_runs";

  if (!supabaseUrl || !serviceRoleKey) {
    return res.status(500).json({ error: "Supabase environment variables are missing." });
  }

  try {
    const payload = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    const companyName = String(payload?.company_name || "").trim();
    if (!companyName) {
      return res.status(400).json({ error: "company_name is required." });
    }

    const response = await fetch(`${supabaseUrl}/rest/v1/${tableName}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "apikey": serviceRoleKey,
        "Authorization": `Bearer ${serviceRoleKey}`,
        "Prefer": "return=representation"
      },
      body: JSON.stringify({
        company_name: companyName,
        uploaded_reports: payload.uploaded_reports || [],
        summary_json: payload.summary_json || {},
        vouchers_json: payload.vouchers_json || [],
        seller_summary_json: payload.seller_summary_json || [],
        gstr1_json: payload.gstr1_json || {},
        gstr3b_json: payload.gstr3b_json || [],
        row_count: Array.isArray(payload.vouchers_json) ? payload.vouchers_json.length : 0,
        created_at: new Date().toISOString()
      })
    });

    const data = await response.json();
    if (!response.ok) {
      console.error("Supabase error:", data);
      return res.status(response.status).json({ error: data });
    }

    return res.status(200).json({ ok: true, id: data?.[0]?.id, data });
  } catch (error) {
    console.error("Handler error:", error);
    return res.status(500).json({ error: error.message || "Unknown error" });
  }
};
