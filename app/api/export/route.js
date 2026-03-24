import { getServerSession } from "next-auth"
import { authOptions } from "../auth/[...nextauth]/route"
import { htmlToText, buildContacts } from "../../../lib/parser"
import ExcelJS from "exceljs"

export const maxDuration = 60

const GRAPH = "https://graph.microsoft.com/v1.0"

async function fetchMails(accessToken, maxMails = 5000) {
  const headers = { Authorization: `Bearer ${accessToken}` }

  // Step 1: Get sender list quickly (no body) - 100 per page
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

  // Step 2: Fetch body only for unique senders (max 200 to stay within timeout)
  const senders = [...senderMap.entries()]
  const toFetch = senders.slice(0, 200)

  const mailItems = await Promise.all(
    toFetch.map(async ([email, info]) => {
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

  // Add remaining senders without body
  for (const [email, info] of senders.slice(200)) {
    mailItems.push({ email, name: info.name, body: "", date: info.date })
  }

  return { mailItems, totalScanned: total }
}

function createExcel(contacts) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet("Kontakte")

  ws.columns = [
    { header: "Vorname",            key: "vorname",   width: 15 },
    { header: "Nachname",           key: "nachname",  width: 20 },
    { header: "Vollst Name",        key: "name",      width: 28 },
    { header: "Firma",              key: "firma",     width: 30 },
    { header: "Position",           key: "position",  width: 32 },
    { header: "Email",              key: "email",     width: 32 },
    { header: "Email 2",            key: "email2",    width: 28 },
    { header: "Telefon",            key: "telefon",   width: 28 },
    { header: "Letzte E-Mail",      key: "date",      width: 16 },
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

  contacts.forEach((contact, i) => {
    const row = ws.addRow(contact)
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
    const contacts = buildContacts(mailItems)
    const wb = createExcel(contacts)
    const buffer = await wb.xlsx.writeBuffer()

    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="outlook_kontakte.xlsx"`,
        "X-Contact-Count": String(contacts.length),
        "X-Mail-Count": String(totalScanned),
      },
    })
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 })
  }
}
