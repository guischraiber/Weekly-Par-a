export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET");

  const { url } = req.query;
  if (!url) return res.status(400).json({ error: "url obrigatória" });

  try {
    const response = await fetch(
      `https://is.gd/create.php?format=simple&url=${encodeURIComponent(url)}`
    );
    const short = await response.text();
    if (!short.startsWith("https://")) throw new Error("Resposta inválida");
    return res.status(200).json({ short });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
