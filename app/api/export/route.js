import { getServerSession } from "next-auth"
import { authOptions } from "../auth/[...nextauth]/route"
import { htmlToText, buildContacts } from "../../../lib/parser"
import ExcelJS from "exceljs"

export const maxDuration = 60

const GRAPH = "https://graph.microsoft.com/v1.0"
const OPENAI_API = "https://api.openai.com/v1/chat/completions"

async function fetchAllSenders(accessToken, maxMails) {
  const headers = { Authorization: `Bearer ${accessToken}` }
  let url = `${GRAPH}/me/mailFolders/inbox/messages?\$top=100&\$select=from,receivedDateTime&\$orderby=receivedDateTime desc`
  const senderMap = new Map()
  let total = 0
  while (url && total < maxMails) {
    const res = await fetch(url, { headers })
    if (!res.ok) break
    const data = await res.json()
    for (const mail of data.value || []) {
      const addr = (mail.from?.emailAddress?.address || "").toLowerCase().trim()
      const name = mail.from?.emailAddress?.name || ""
      const date = (mail.receivedDateTime || "").slice(0, 10)
      if (!addr) continue
      if (!senderMap.has(addr)) senderMap.set(addr, { name, date, id: mail.id })
      total++
    }
    url = data["@odata.nextLink"] || null
  }
  return { senderMap, total }
}

async function fetchBodies(accessToken, senders) {
  const headers = { Authorization: `Bearer ${accessToken}` }
  const BATCH = 20
  const mailItems = []
  for (let i = 0; i < senders.length; i += BATCH) {
    const batch = senders.slice(i, i + BATCH)
    const results = await Promise.all(batch.map(async ([email, info]) => {
      try {
        const res = await fetch(`${GRAPH}/me/messages/${info.id}?\$select=uniqueBody`, { headers })
        if (!res.ok) return { email, name: info.name, body: "", date: info.date }
        const data = await res.json()
        const raw = data.uniqueBody?.content || ""
        const body = data.uniqueBody?.contentType === "html" ? htmlToText(raw) : raw
        return { email, name: info.name, body, date: info.date }
      } catch { return { email, name: info.name, body: "", date: info.date } }
    }))
    mailItems.push(...results)
  }
  return mailItems
}

function extractSig(body) {
  // Give GPT the full mail body (last 60 lines, stripped of noise)
  // Position can be anywhere: above/below signature, after greeting, etc.
  const lines = body.split("\n")
    .map(l => l.trim())
    .filter(l => l.length > 0 && l.length < 200)
    .filter(l => !/^(https?:\/\/|www\.|ARC-|DKIM-|Content-|Received:|Message-ID|X-MS|Thread-|Return-Path)/i.test(l))
  return lines.slice(-60).join("\n")
}

async function gptEnrich(contacts) {
  if (!process.env.OPENAI_API_KEY) return contacts
  const toEnrich = contacts.filter(c => c._sig && c._sig.length > 10)
  if (!toEnrich.length) return contacts

  const BATCH = 40
  for (let i = 0; i < toEnrich.length; i += BATCH) {
    const batch = toEnrich.slice(i, i + BATCH)
    const prompt = `Du bekommst E-Mail-Inhalte von Kontakten. Extrahiere Jobtitel/Position und Firmenname.
Antworte NUR mit einem JSON-Array, kein anderer Text:
[{"id":0,"position":"Assistenz der Geschaeftsfuehrung","firma":"Feser Graf GmbH"},{"id":1,"position":"","firma":""}]

Regeln:
- position = Berufsbezeichnung/Jobtitel (ALLE Berufe, auch einfache wie Sachbearbeiter, Assistent, Techniker etc.)
- Position steht oft direkt unter dem Namen, nach Grussformel, oder in der Signatur
- firma = vollstaendiger offizieller Firmenname
- Leerer String wenn nicht erkennbar

${batch.map((c, idx) => `[ID:${idx}] Name: ${c.name}\n${c._sig}`).join("\n===\n")}`

    try {
      const res = await fetch(OPENAI_API, {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${process.env.OPENAI_API_KEY}` },
        body: JSON.stringify({ model: "gpt-4o-mini", max_tokens: 2000, messages: [{ role: "user", content: prompt }] }),
      })
      if (!res.ok) continue
      const data = await res.json()
      const text = data.choices?.[0]?.message?.content || ""
      const match = text.match(/\[[\s\S]*?\]/)
      if (!match) continue
      const results = JSON.parse(match[0])
      for (const r of results) {
        if (typeof r.id === 'number' && r.id >= 0 && r.id < batch.length) {
          if (r.position) batch[r.id].position = r.position
          if (r.firma && !batch[r.id].firma) batch[r.id].firma = r.firma
        }
      }
    } catch (e) {}
  }
  return contacts
}

function createExcel(contacts) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet("Kontakte")
  ws.columns = [
    { header: "Vorname", key: "vorname", width: 15 },
    { header: "Nachname", key: "nachname", width: 20 },
    { header: "Name", key: "name", width: 28 },
    { header: "Firma", key: "firma", width: 30 },
    { header: "Position", key: "position", width: 32 },
    { header: "Email", key: "email", width: 32 },
    { header: "Email 2", key: "email2", width: 28 },
    { header: "Telefon", key: "telefon", width: 28 },
    { header: "Letzte E-Mail", key: "date", width: 16 },
  ]
  ws.getRow(1).eachCell(cell => {
    cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" }, size: 10 }
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } }
    cell.alignment = { horizontal: "center", vertical: "middle" }
    cell.border = { top:{style:"thin",color:{argb:"FFB0B0B0"}},bottom:{style:"thin",color:{argb:"FFB0B0B0"}},left:{style:"thin",color:{argb:"FFB0B0B0"}},right:{style:"thin",color:{argb:"FFB0B0B0"}} }
  })
  ws.getRow(1).height = 28
  ws.views = [{ state: "frozen", ySplit: 1 }]
  ws.autoFilter = { from: "A1", to: "I1" }
  contacts.forEach((c, i) => {
    const row = ws.addRow(c)
    const fill = i % 2 === 0 ? { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCE6F1" } } : undefined
    row.eachCell(cell => {
      cell.font = { name: "Arial", size: 9 }
      cell.alignment = { vertical: "top" }
      cell.border = { top:{style:"thin",color:{argb:"FFB0B0B0"}},bottom:{style:"thin",color:{argb:"FFB0B0B0"}},left:{style:"thin",color:{argb:"FFB0B0B0"}},right:{style:"thin",color:{argb:"FFB0B0B0"}} }
      if (fill) cell.fill = fill
    })
  })
  return wb
}

export async function GET(request) {
  const session = await getServerSession(authOptions)
  if (!session?.accessToken) {
    return new Response(JSON.stringify({ error: "Nicht eingeloggt" }), { status: 401 })
  }
  const { searchParams } = new URL(request.url)
  const maxMails = parseInt(searchParams.get("max") || "5000")

  try {
    // Step 1: Get all sender IDs fast
    const { senderMap, total } = await fetchAllSenders(session.accessToken, maxMails)
    const senders = [...senderMap.entries()]

    // Step 2: Fetch bodies for all unique senders
    const mailItems = await fetchBodies(session.accessToken, senders)

    // Step 3: Build contacts with regex parser
    let contacts = buildContacts(mailItems)

    // Step 4: Build signature map
    const sigMap = new Map()
    for (const item of mailItems) {
      if (item.body && !sigMap.has(item.email)) {
        sigMap.set(item.email, extractSig(item.body))
      }
    }

    // Step 5: Attach sigs and enrich with GPT in parallel
    contacts = contacts.map(c => ({ ...c, _sig: sigMap.get(c.email) || "" }))
    contacts = await gptEnrich(contacts)

    // Clean output
    const out = contacts.map(c => ({
      vorname: c.vorname, nachname: c.nachname, name: c.name,
      firma: c.firma, position: c.position,
      email: c.email, email2: c.email2,
      telefon: c.telefon, date: c.date,
    }))

    const wb = createExcel(out)
    const buffer = await wb.xlsx.writeBuffer()

    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="outlook_kontakte.xlsx"`,
        "X-Contact-Count": String(out.length),
        "X-Mail-Count": String(total),
      },
    })
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 })
  }
}