import { getServerSession } from "next-auth"
import { authOptions } from "../auth/[...nextauth]/route"
import { htmlToText, buildContacts } from "../../../lib/parser"
import ExcelJS from "exceljs"

export const maxDuration = 60

const GRAPH = "https://graph.microsoft.com/v1.0"
const OPENAI_API = "https://api.openai.com/v1/chat/completions"

async function fetchMails(accessToken, maxMails = 5000) {
  const headers = { Authorization: `Bearer ${accessToken}` }
  let url = `${GRAPH}/me/mailFolders/inbox/messages?$top=100&$select=from,receivedDateTime&$orderby=receivedDateTime desc`
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
      if (!senderMap.has(addr)) {
        senderMap.set(addr, { name, date, id: mail.id })
      }
      total++
    }
    url = data["@odata.nextLink"] || null
  }

  const senders = [...senderMap.entries()]

  // Fetch body for all unique senders in batches of 20
  const BATCH = 20
  const mailItems = []

  for (let i = 0; i < senders.length; i += BATCH) {
    const batch = senders.slice(i, i + BATCH)
    const results = await Promise.all(
      batch.map(async ([email, info]) => {
        try {
          const res = await fetch(
            `${GRAPH}/me/messages/${info.id}?$select=uniqueBody`,
            { headers }
          )
          if (!res.ok) return { email, name: info.name, body: "", date: info.date }
          const data = await res.json()
          const raw = data.uniqueBody?.content || ""
          const body = data.uniqueBody?.contentType === "html" ? htmlToText(raw) : raw
          return { email, name: info.name, body, date: info.date }
        } catch {
          return { email, name: info.name, body: "", date: info.date }
        }
      })
    )
    mailItems.push(...results)
  }

  return { mailItems, totalScanned: total }
}

// Extract signature lines from body text
function extractSignatureLines(body) {
  const lines = body.split("\n")
  // Find signature separator
  for (let i = lines.length - 1; i > Math.max(0, lines.length - 120); i--) {
    const t = lines[i].trim()
    if (/^(_{2,}|-{2,})|mit freundlichen|with kind regards|best regards|freundliche gr|freundlichen gr|viele gr|herzliche gr|kind regards|sincerely|^mfg\s*$|^lg[,.]?\s*$|^vg\s*$/i.test(t)) {
      return lines.slice(i, i + 40).map(l => l.trim()).filter(l => l && l.length < 120).join("\n")
    }
  }
  return lines.slice(-35).map(l => l.trim()).filter(l => l && l.length < 120).join("\n")
}

// Use Claude AI to extract position and company from signature
async function extractWithAI(contacts) {
  // Only call AI for contacts that have a body/signature
  const needsAI = contacts.filter(c => c._sig && c._sig.length > 20)
  if (!needsAI.length || !process.env.OPENAI_API_KEY) return contacts

  // Build batch prompt for all contacts at once (max 50 per call)
  const AIBATCH = 50
  for (let i = 0; i < needsAI.length; i += AIBATCH) {
    const batch = needsAI.slice(i, i + AIBATCH)
    const prompt = `Du bekommst E-Mail-Signaturen von Kontakten. Extrahiere fuer jeden Kontakt die Berufsposition und Firmenname.

Antworte NUR mit einem JSON-Array in diesem Format (keine anderen Texte):
[{"id":0,"position":"Verkaufsberater","firma":"BMW Group"},{"id":1,"position":"","firma":""}]

Regeln:
- "position" = Berufsbezeichnung/Jobtitel (z.B. "Geschaeftsfuehrer", "Verkaufsberater", "Marketing Manager", "Assistentin", "Techniker")
- "firma" = Firmenname (nur wenn klar erkennbar, sonst leer)
- Falls keine Position erkennbar: leerer String ""
- Kuerze Positionen auf das Wesentliche
- Erkenne alle Berufe, nicht nur Fuehrungskraefte

Signaturen:
${batch.map((c, idx) => `[${idx}] Name: ${c.name}\n${c._sig}`).join("\n\n---\n\n")}
`

    try {
      const res = await fetch(OPENAI_API, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${process.env.OPENAI_API_KEY}`,
        },
        body: JSON.stringify({
          model: "gpt-4o-mini",
          max_tokens: 2000,
          messages: [{ role: "user", content: prompt }],
        }),
      })

      if (!res.ok) continue
      const data = await res.json()
      const text = data.choices?.[0]?.message?.content || ""
      const jsonMatch = text.match(/\[[\s\S]*\]/)
      if (!jsonMatch) continue

      const results = JSON.parse(jsonMatch[0])
      for (const r of results) {
        if (r.id >= 0 && r.id < batch.length) {
          const contact = batch[r.id]
          if (r.position && !contact.position) contact.position = r.position
          if (r.firma && !contact.firma) contact.firma = r.firma
        }
      }
    } catch {
      // Continue without AI for this batch
    }
  }

  return contacts
}

function createExcel(contacts) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet("Kontakte")

  ws.columns = [
    { header: "Vorname",       key: "vorname",   width: 15 },
    { header: "Nachname",      key: "nachname",  width: 20 },
    { header: "Name",          key: "name",      width: 28 },
    { header: "Firma",         key: "firma",     width: 30 },
    { header: "Position",      key: "position",  width: 32 },
    { header: "Email",         key: "email",     width: 32 },
    { header: "Email 2",       key: "email2",    width: 28 },
    { header: "Telefon",       key: "telefon",   width: 28 },
    { header: "Letzte E-Mail", key: "date",      width: 16 },
  ]

  ws.getRow(1).eachCell(cell => {
    cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" }, size: 10 }
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } }
    cell.alignment = { horizontal: "center", vertical: "middle" }
    cell.border = {
      top: { style: "thin", color: { argb: "FFB0B0B0" } },
      bottom: { style: "thin", color: { argb: "FFB0B0B0" } },
      left: { style: "thin", color: { argb: "FFB0B0B0" } },
      right: { style: "thin", color: { argb: "FFB0B0B0" } },
    }
  })
  ws.getRow(1).height = 28
  ws.views = [{ state: "frozen", ySplit: 1 }]
  ws.autoFilter = { from: "A1", to: "I1" }

  contacts.forEach((c, i) => {
    const row = ws.addRow(c)
    const fill = i % 2 === 0
      ? { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCE6F1" } }
      : undefined
    row.eachCell(cell => {
      cell.font = { name: "Arial", size: 9 }
      cell.alignment = { vertical: "top" }
      cell.border = {
        top: { style: "thin", color: { argb: "FFB0B0B0" } },
        bottom: { style: "thin", color: { argb: "FFB0B0B0" } },
        left: { style: "thin", color: { argb: "FFB0B0B0" } },
        right: { style: "thin", color: { argb: "FFB0B0B0" } },
      }
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
    const { mailItems, totalScanned } = await fetchMails(session.accessToken, maxMails)
    
    // Build basic contacts first
    let contacts = buildContacts(mailItems)
    
    // Add signature text to each contact for AI processing
    const sigMap = new Map()
    for (const item of mailItems) {
      if (item.body && !sigMap.has(item.email)) {
        sigMap.set(item.email, extractSignatureLines(item.body))
      }
    }
    
    // Attach signature and prepare for AI
    contacts = contacts.map(c => ({
      ...c,
      _sig: sigMap.get(c.email) || sigMap.get(c.email2) || "",
    }))
    
    // Use Claude AI to extract positions
    contacts = await extractWithAI(contacts)
    
    // Remove internal fields before Excel
    const cleanContacts = contacts.map(({ _sig, _key, _email, _freemail, _all, ...rest }) => ({
      vorname: rest.vorname,
      nachname: rest.nachname,
      name: rest.name,
      firma: rest.firma,
      position: rest.position,
      email: rest.email,
      email2: rest.email2,
      telefon: rest.telefon,
      date: rest.date,
    }))

    const wb = createExcel(cleanContacts)
    const buffer = await wb.xlsx.writeBuffer()

    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="outlook_kontakte.xlsx"`,
        "X-Contact-Count": String(cleanContacts.length),
        "X-Mail-Count": String(totalScanned),
      },
    })
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 })
  }
}
