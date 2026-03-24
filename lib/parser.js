// Filter patterns for spam/automated emails
const FILTER_PATTERNS = [
  /noreply/i, /no-reply/i, /donotreply/i, /do-not-reply/i,
  /mailer-daemon/i, /postmaster/i, /bounce/i,
  /newsletter/i, /mailing/i, /campaign/i,
  /notification/i, /alert@/i, /automated/i,
  /@mailchimp/i, /@sendgrid/i, /@brevo/i, /@klaviyo/i,
  /@hubspot/i, /@amazonses/i, /billing@/i, /invoice@/i,
  /no.reply/i, /unsubscribe/i, /@em\d+\./i,
]

const FREEMAIL = new Set([
  "gmail.com","googlemail.com","yahoo.com","yahoo.de","yahoo.at",
  "hotmail.com","hotmail.de","outlook.com","live.com","live.de",
  "icloud.com","me.com","gmx.at","gmx.de","gmx.net","web.de",
  "aon.at","chello.at","protonmail.com","pm.me","proton.me","t-online.de",
])

const QUOTE_RE = /^(-{4,}|_{4,}|Von:\s|From:\s|Am\s.{5,60}schrieb|On\s.{5,60}wrote:|-----\s*Original|\[mailto:)/im

const SIG_SEP_RE = /^(_{2,}|-{2,})|mit freundlichen|with kind regards|best regards|freundliche grüße|freundlichen grüßen|viele grüße|herzliche grüße|kind regards|sincerely|^mfg\s*$|^lg[,.]?\s*$|^vg\s*$/im

const PHONE_RE = /(?:(?:tel|fon|phone|mob|mobil|t|m|d|fax|direct|büro|office)[\s:./\-]*)?(\+?[\d][\d\s\(\)\-\/\.]{5,18}\d)/gi

const POSITION_KEYWORDS = [
  "manager","managerin","director","leiter","leiterin","geschäftsführer",
  "ceo","cfo","cto","coo","vp ","vice president","head of","chief",
  "berater","beraterin","consultant","analyst","spezialist",
  "koordinator","assistent","ingenieur","entwickler",
  "projektleiter","teamleiter","abteilungsleiter",
  "officer","partner","associate","referent","sachbearbeiter",
  "kaufmann","kauffrau","vertrieb","sales","marketing","finance",
  "rechtsanwalt","steuerberater","wirtschaftsprüfer",
  "prokurist","vorstand","founder","owner","inhaber",
  "außendienst","innendienst","kundenberater","account manager",
  "gebietsleiter","selbstständig","freelance",
]

const COMPANY_RE = /\b(GmbH|AG|KG|OG|GbR|Ltd\.?|Inc\.?|Corp\.?|SE\b|e\.U\.|KEG|GmbH\s*&\s*Co\.?\s*KG|Gruppe\b|Group\b|Holding\b|Stiftung\b|Consulting\b|Solutions\b|Services\b|GesmbH)\b/i

const NOISE_RE = /www\.|https?:\/\/|\bimpress|\bdatenschutz|\bdisclaimer|\bvertraulich\b|\bconfidential\b|\bunsubscribe\b|\bcopyright\b|\bregistergericht\b|\bsteuer.?nr\b|\biban\b|\bbic\b|\bdiese.*e.?mail\b|\bagb\b/i

export function isFiltered(email) {
  return FILTER_PATTERNS.some(p => p.test(email))
}

export function isFreemail(email) {
  const domain = email.split("@")[1]?.toLowerCase() || ""
  return FREEMAIL.has(domain)
}

export function looksLikeEmail(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s.trim())
}

export function cleanName(name, email) {
  name = (name || "").trim()
  if (!name || looksLikeEmail(name) || name.length < 2) {
    const local = email.split("@")[0] || ""
    const parts = local.split(/[._\-]/).filter(p => p.length > 1)
    return parts.slice(0, 2).map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(" ") || email
  }
  if (name.includes(",") && name.split(",").length === 2) {
    const [last, first] = name.split(",").map(p => p.trim())
    if (last.length > 1 && first.length > 1) return `${first} ${last}`
  }
  return name
}

export function stripQuoted(text) {
  const match = QUOTE_RE.exec(text)
  return match ? text.slice(0, match.index) : text
}

export function extractSignature(text, maxLines = 35) {
  const lines = text.split("\n")
  for (let i = lines.length - 1; i > Math.max(0, lines.length - 120); i--) {
    if (SIG_SEP_RE.test(lines[i].trim())) {
      return lines.slice(i).join("\n")
    }
  }
  return lines.slice(-maxLines).join("\n")
}

export function validatePhone(num) {
  const digits = num.replace(/\D/g, "")
  if (digits.length < 6 || digits.length > 15) return null
  if (digits.length <= 6 && !num.trim().startsWith("+")) return null
  if (/^(19|20)\d{2}$/.test(digits)) return null
  return num.trim()
}

export function parseSignature(sig, senderName = "") {
  const phones = []
  let position = ""
  let firma = ""

  const lines = sig.split("\n").map(l => l.trim()).filter(Boolean)

  for (const line of lines) {
    if (NOISE_RE.test(line) || line.length > 120) continue

    let phoneMatch
    const phoneRe = new RegExp(PHONE_RE.source, "gi")
    while ((phoneMatch = phoneRe.exec(line)) !== null) {
      const num = validatePhone(phoneMatch[1])
      if (num && !phones.includes(num)) phones.push(num)
    }

    const lineClean = line.replace(PHONE_RE, "").replace(/[|–\-:,./()[\]{}\s*]+$/, "").trim()
    if (!lineClean || lineClean.length < 3) continue

    if (!position && POSITION_KEYWORDS.some(kw => lineClean.toLowerCase().includes(kw))) {
      if (!looksLikeEmail(lineClean) && lineClean.length < 80) {
        position = lineClean
      }
    }

    if (!firma && COMPANY_RE.test(lineClean)) {
      if (!senderName.toLowerCase().includes(lineClean.toLowerCase().slice(0, 5))) {
        if (!looksLikeEmail(lineClean) && lineClean.length < 80) {
          firma = lineClean
        }
      }
    }
  }

  return { position, phones: phones.slice(0, 4).join(" | "), firma }
}

export function getDomainFirma(email) {
  const domain = email.split("@")[1]?.toLowerCase() || ""
  if (FREEMAIL.has(domain)) return ""
  const parts = domain.split(".")
  if (parts.length >= 2) {
    const name = parts[parts.length - 2]
    if (["info","mail","email","contact","office","support","service"].includes(name)) return ""
    return name.charAt(0).toUpperCase() + name.slice(1)
  }
  return ""
}

export function htmlToText(html) {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<\/td>/gi, " ")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
}

export function buildContacts(mailItems) {
  const senderMap = new Map()

  for (const item of mailItems) {
    const email = item.email.toLowerCase().trim()
    const name = item.name.trim()
    if (!email || isFiltered(email)) continue

    const cleanBody = stripQuoted(item.body)

    if (!senderMap.has(email)) {
      senderMap.set(email, { name, body: cleanBody, date: item.date })
    } else {
      const ex = senderMap.get(email)
      if (item.date > ex.date) {
        ex.body = cleanBody
        ex.date = item.date
      }
      if (!looksLikeEmail(name) && looksLikeEmail(ex.name)) {
        ex.name = name
      }
    }
  }

  const parsed = []
  for (const [email, data] of senderMap) {
    const name = cleanName(data.name, email)
    const sig = extractSignature(data.body)
    const { position, phones, firma: sigFirma } = parseSignature(sig, name)
    const firma = sigFirma || getDomainFirma(email)

    parsed.push({
      _key: name.toLowerCase().trim().replace(/\s+/g, " "),
      _email: email,
      _freemail: isFreemail(email),
      name, firma, position,
      email, email2: "",
      phones,
      date: data.date,
    })
  }

  const GENERIC = new Set(["info","service","office","support","team","kontakt","contact","sales","marketing"])
  const merged = new Map()

  for (const c of parsed) {
    let key = c._key
    if (key.length < 4 || GENERIC.has(key.split(" ")[0])) key = c._email

    if (!merged.has(key)) {
      merged.set(key, { ...c, _all: [c._email] })
    } else {
      const ex = merged.get(key)
      ex._all.push(c._email)
      if (!ex.firma && c.firma) ex.firma = c.firma
      if (!ex.position && c.position) ex.position = c.position
      if (!ex.phones && c.phones) ex.phones = c.phones
      if (c.date > ex.date) ex.date = c.date
      if (ex._freemail && !c._freemail) { ex.email = c._email; ex._freemail = false }
    }
  }

  const result = []
  for (const c of merged.values()) {
    const allEmails = [...new Set(c._all)]
    const parts = c.name.trim().split(" ")
    result.push({
      vorname: parts[0] || "",
      nachname: parts.slice(1).join(" ") || "",
      name: c.name,
      firma: c.firma,
      position: c.position,
      email: allEmails[0],
      email2: allEmails[1] || "",
      telefon: c.phones,
      date: c.date,
    })
  }

  return result.sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()))
}
