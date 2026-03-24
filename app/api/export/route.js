import { getServerSession } from "next-auth"
import { authOptions } from "../auth/[...nextauth]/route"
import { htmlToText, buildContacts } from "../../../lib/parser"
import ExcelJS from "exceljs"

const GRAPH = "https://graph.microsoft.com/v1.0"

async function fetchMails(accessToken, maxMails = 5000) {
  const headers = { Authorization: `Bearer ${accessToken}` }
  let url = `${GRAPH}/me/mailFolders/inbox/messages?$top=100&$select=from,receivedDateTime,body&$orderby=receivedDateTime desc`
  const mailItems = []

  while (url && mailItems.length < maxMails) {
    const res = await fetch(url, { headers })
    if (!res.ok) break
    const data = await res.json()

    for (const mail of data.value || []) {
      try {
        const sender = mail.from?.emailAddress || {}
        const bodyObj = mail.body || {}
        const raw = bodyObj.content || ""
        const body = bodyObj.contentType === "html" ? htmlToText(raw) : raw
        mailItems.push({
          email: (sender.address || "").toLowerCase().trim(),
          name: sender.name || "",
          body,
          date: (mail.receivedDateTime || "").slice(0, 10),
        })
      } catch {}
    }

    url = data["@odata.nextLink"] || null
    await new Promise(r => setTimeout(r, 30))
  }

  return mailItems
}

function createExcel(contacts) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet("Kontakte")

  const COLUMNS = [
    { header: "Vorname",            key: "vorname",   width: 15 },
    { header: "Nachname",           key: "nachname",  width: 20 },
    { header: "Vollständiger Name", key: "name",      width: 28 },
    { header: "Firma",              key: "firma",     width: 30 },
    { header: "Position",           key: "position",  width: 32 },
    { header: "Email",              key: "email",     width: 32 },
    { header: "Email 2",            key: "email2",    width: 28 },
    { header: "Telefon",            key: "telefon",   width: 28 },
    { header: "Letzte E-Mail",      key: "date",      width: 16 },
  ]

  ws.columns = COLUMNS

  ws.getRow(1).eachCell(cell => {
    cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" }, size: 10 }
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } }
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true }
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
    const mailItems = await fetchMails(session.accessToken, maxMails)
    const contacts = buildContacts(mailItems)
    const wb = createExcel(contacts)

    const buffer = await wb.xlsx.writeBuffer()

    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="outlook_kontakte.xlsx"`,
        "X-Contact-Count": String(contacts.length),
        "X-Mail-Count": String(mailItems.length),
      },
    })
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 })
  }
}
