module.exports = async function handler(req, res) {
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
    const response = await fetch(`${supabaseUrl}/rest/v1/${tableName}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "apikey": serviceRoleKey,
        "Authorization": `Bearer ${serviceRoleKey}`,
        "Prefer": "return=representation"
      },
      body: JSON.stringify({
        company_name: payload.company_name || "",
        uploaded_reports: payload.uploaded_reports || [],
        summary_json: payload.summary_json || {},
        vouchers_json: payload.vouchers_json || [],
        seller_summary_json: payload.seller_summary_json || [],
        gstr1_json: payload.gstr1_json || {},
        gstr3b_json: payload.gstr3b_json || [],
        created_at: new Date().toISOString()
      })
    });

    const data = await response.json();
    if (!response.ok) {
      return res.status(response.status).json({ error: data });
    }

    return res.status(200).json({ ok: true, data });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Unknown error" });
  }
};
