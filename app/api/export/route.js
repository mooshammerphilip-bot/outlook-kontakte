import { getServerSession } from "next-auth"
import { authOptions } from "../auth/[...nextauth]/route"
import { htmlToText, buildContacts } from "../../../lib/parser"
import ExcelJS from "exceljs"

export const maxDuration = 60

const GRAPH = "https://graph.microsoft.com/v1.0"

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

// Find position and phone directly from mail body
// Searches both near the sender name AND in the signature block
function extractPositionAndPhone(body, senderName) {
  const lines = body.split("\n").map(l => l.trim()).filter(l => l.length > 0 && l.length < 150)
  
  let position = ""
  let phone = ""
  let firma = ""
  
  const NOISE = /www\.|https?:\/\/|@|registergericht|steuer|iban|bic|datenschutz|disclaimer|confidential|unsubscribe|copyright|agb|fax:|\bfax\b/i
  const PHONE_RE = /(?:(?:tel|fon|phone|mob|mobil|t|m|d|direct)[s:./\-]*)?(\ *\+?[0-9][0-9\s()\-\/\.]{5,18}[0-9])/i
  const COMPANY_RE = /\b(GmbH|AG|KG|OG|GbR|Ltd\.?|Inc\.?|Corp\.?|SE\b|e\.U\.|KEG|GmbH\s*&\s*Co\.?\s*KG|Gruppe\b|Group\b|Holding\b|GesmbH|Automobil|Autohaus|Fahrzeug|Werkstatt)\b/i
  
  // Position patterns - covers all job types
  const POS_RE = /\b(geschaeftsfuehrer|geschaeftsfuehrerin|geschaeftsfuehrungi?\.?a\.|i\s*\.?\s*a\.|geschaftsfuhrung|geschaftsleit|geschaftsleiter|inhaber|eigentuemer|owner|founder|ceo|cfo|cto|coo|cso|vorstand|direktor|director|prokurist|leiter|leiterin|head of|vp |vice pres|managing|partner|gesellschafter|verkauf|verkaeufer|verkaufsberater|vertrieb|account manager|sales|aussendienst|innendienst|kundenberater|berater|beraterin|consultant|advisor|marketing|pr-?manager|redakteur|content|techniker|ingenieur|engineer|entwickler|developer|programmierer|it-?leiter|administrator|personal|hr |recruiter|assistent|assistentin|assistant|office manager|sachbearbeiter|referent|sekretaer|projektleiter|projektmanager|koordinator|finanz|controller|serviceleiter|servicetechniker|werkstattleiter|after.?sales|kundendienst|support)\b/i
  
  // Strategy 1: Look for position NEAR the sender name in the mail
  const nameParts = senderName.toLowerCase().split(/\s+/).filter(p => p.length > 2)
  let nameLineIdx = -1
  
  for (let i = 0; i < lines.length; i++) {
    const lineLow = lines[i].toLowerCase()
    if (nameParts.length >= 2 && nameParts.every(p => lineLow.includes(p))) {
      nameLineIdx = i
      break
    }
    if (nameParts.length >= 1 && nameParts[0].length > 3 && lineLow.includes(nameParts[0]) && 
        (nameParts.length === 1 || lineLow.includes(nameParts[nameParts.length-1]))) {
      nameLineIdx = i
      break
    }
  }
  
  // Check 5 lines after the name for position
  if (nameLineIdx >= 0) {
    for (let i = nameLineIdx + 1; i < Math.min(nameLineIdx + 6, lines.length); i++) {
      const line = lines[i]
      if (!line || NOISE.test(line)) continue
      if (line.length > 80) continue
      if (!position && POS_RE.test(line)) {
        position = line
      }
      if (!firma && COMPANY_RE.test(line)) {
        firma = line
      }
    }
  }
  
  // Strategy 2: Find signature block and extract from there
  let sigStart = -1
  const GREETING_RE = /^(mit (freundlichen?|besten?|herzlichen?|lieben?|kollegialen?|vielen?)|with (kind|best|warm) regards|best regards|kind regards|viele gr|herzlich|freundlich|sincerely|mit gr|gruss|gruesse|^mfg|^vg\s*$|^lg[,.]?\s*$|^ciao\s*$)/i
  
  for (let i = lines.length - 1; i > Math.max(0, lines.length - 80); i--) {
    if (GREETING_RE.test(lines[i]) || /^(_{2,}|-{2,})/.test(lines[i])) {
      sigStart = i
      break
    }
  }
  
  const sigLines = sigStart >= 0 ? lines.slice(sigStart) : lines.slice(-35)
  
  for (const line of sigLines) {
    if (NOISE.test(line) || line.length > 120) continue
    
    // Extract phone
    if (!phone) {
      const m = PHONE_RE.exec(line)
      if (m) {
        const digits = m[1].replace(/\D/g, "")
        if (digits.length >= 7 && digits.length <= 15 && !/^(19|20)\d{2}$/.test(digits)) {
          phone = m[1].trim()
        }
      }
    }
    
    // Extract position from sig if not found near name
    if (!position && POS_RE.test(line) && line.length < 80) {
      position = line.replace(PHONE_RE, "").trim()
    }
    
    // Extract firma from sig
    if (!firma && COMPANY_RE.test(line) && line.length < 100) {
      firma = line.replace(PHONE_RE, "").trim()
    }
  }
  
  return { position, phone, firma }
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
    const { senderMap, total } = await fetchAllSenders(session.accessToken, maxMails)
    const senders = [...senderMap.entries()]
    const mailItems = await fetchBodies(session.accessToken, senders)
    
    // Build base contacts with parser (phone + basic firma)
    let contacts = buildContacts(mailItems)
    
    // Build body map for enhanced extraction
    const bodyMap = new Map()
    for (const item of mailItems) {
      if (item.body && !bodyMap.has(item.email)) bodyMap.set(item.email, item.body)
    }
    
    // Enhance each contact with smart position/phone extraction
    contacts = contacts.map(c => {
      const body = bodyMap.get(c.email) || bodyMap.get(c.email2) || ""
      if (!body) return c
      const { position, phone, firma } = extractPositionAndPhone(body, c.name)
      return {
        ...c,
        position: c.position || position,
        telefon: c.telefon || phone,
        firma: c.firma || firma,
      }
    })

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