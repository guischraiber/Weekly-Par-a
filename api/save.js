import { kv } from "@vercel/kv";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).end();

  try {
    const { data } = req.body;
    if (!data) return res.status(400).json({ error: "data obrigatório" });

    // Gerar ID curto de 6 chars
    const id = Math.random().toString(36).substring(2, 8);

    // Salvar por 30 dias
    await kv.set(`w:${id}`, data, { ex: 60 * 60 * 24 * 30 });

    return res.status(200).json({ id });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
