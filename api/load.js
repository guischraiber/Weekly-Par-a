import { kv } from "@vercel/kv";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");

  const { id } = req.query;
  if (!id) return res.status(400).json({ error: "id obrigatório" });

  try {
    const data = await kv.get(`w:${id}`);
    if (!data) return res.status(404).json({ error: "não encontrado ou expirado" });
    return res.status(200).json({ data });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
