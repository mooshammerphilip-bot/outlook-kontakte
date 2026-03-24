"use client"
import { useSession, signIn, signOut } from "next-auth/react"
import { useState } from "react"

export default function Home() {
  const { data: session, status } = useSession()
  const [loading, setLoading] = useState(false)
  const [progress, setProgress] = useState("")
  const [result, setResult] = useState(null)
  const [maxMails, setMaxMails] = useState(500)
  const [error, setError] = useState("")

  async function handleExport() {
    setLoading(true)
    setError("")
    setResult(null)
    setProgress("Verbinde mit Microsoft...")
    try {
      setProgress(`Lade ${maxMails.toLocaleString()} Mails aus Posteingang...`)
      const res = await fetch(`/api/export?max=${maxMails}`)
      if (!res.ok) {
        const err = await res.json()
        throw new Error(err.error || "Export fehlgeschlagen")
      }
            const count = res.headers.get("X-Contact-Count")
      const mails = res.headers.get("X-Mail-Count")
      setProgress("Erstelle Excel-Datei...")
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement("a")
      a.href = url
      a.download = "outlook_kontakte.xlsx"
      a.click()
      URL.revokeObjectURL(url)
      setResult({ contacts: count, mails })
      setProgress("")
    } catch (e) {
      setError(e.message)
      setProgress("")
    } finally {
      setLoading(false)
    }
  }

  return (
    <main style={{
      minHeight: "100vh",
      background: "#F0F4F8",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontFamily: "Arial, sans-serif",
      padding: "20px",
    }}>
      <div style={{
        background: "white",
        borderRadius: "16px",
        boxShadow: "0 4px 24px rgba(0,0,0,0.08)",
        width: "100%",
        maxWidth: "480px",
        overflow: "hidden",
      }}>
        <div style={{
          background: "#1F4E79",
          padding: "28px 32px",
          textAlign: "center",
        }}>
          <h1 style={{ color: "white", margin: 0, fontSize: "20px", fontWeight: "700" }}>
            Outlook Kontakte Exporter
          </h1>
          <p style={{ color: "#A8C8E8", margin: "6px 0 0", fontSize: "13px" }}>
            Kontakte aus E-Mail-Signaturen extrahieren
          </p>
        </div>

        <div style={{ padding: "28px 32px" }}>
          {status === "loading" && (
            <div style={{ textAlign: "center", color: "#888", padding: "20px 0" }}>
              Ladt...
            </div>
          )}

          {status === "unauthenticated" && (
            <div style={{ textAlign: "center" }}>
              <p style={{ color: "#555", marginBottom: "20px", fontSize: "14px" }}>
                Bitte mit Microsoft-Konto anmelden.
              </p>
              <button
                onClick={() => signIn("azure-ad")}
                style={{
                  background: "#0078D4",
                  color: "white",
                  border: "none",
                  borderRadius: "8px",
                  padding: "14px 32px",
                  fontSize: "15px",
                  fontWeight: "600",
                  cursor: "pointer",
                  width: "100%",
                }}
              >
                Mit Microsoft einloggen
              </button>
            </div>
          )}

          {status === "authenticated" && (
            <div>
              <div style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                background: "#F0FFF4",
                border: "1px solid #C6F6D5",
                borderRadius: "8px",
                padding: "10px 14px",
                marginBottom: "20px",
              }}>
                <div>
                  <div style={{ fontSize: "13px", fontWeight: "600", color: "#276749" }}>
                    Eingeloggt
                  </div>
                  <div style={{ fontSize: "11px", color: "#48BB78" }}>
                    {session.user?.email}
                  </div>
                </div>
                <button
                  onClick={() => signOut()}
                  style={{
                    background: "none",
                    border: "1px solid #C6F6D5",
                    borderRadius: "6px",
                    padding: "4px 10px",
                    fontSize: "11px",
                    color: "#276749",
                    cursor: "pointer",
                  }}
                >
                  Abmelden
                </button>
              </div>

              <div style={{ marginBottom: "20px" }}>
                <label style={{
                  display: "block",
                  fontSize: "13px",
                  color: "#555",
                  marginBottom: "8px",
                  fontWeight: "600",
                }}>
                  Anzahl Mails scannen:
                </label>
                <div style={{ display: "flex", gap: "8px" }}>
                  {[1000, 2000, 5000, 10000].map(n => (
                    <button
                      key={n}
                      onClick={() => setMaxMails(n)}
                      style={{
                        flex: 1,
                        padding: "8px 4px",
                        borderRadius: "8px",
                        border: maxMails === n ? "2px solid #1F4E79" : "1px solid #DDD",
                        background: maxMails === n ? "#EBF3FB" : "white",
                        color: maxMails === n ? "#1F4E79" : "#555",
                        fontWeight: maxMails === n ? "700" : "400",
                        fontSize: "13px",
                        cursor: "pointer",
                      }}
                    >
                      {n.toLocaleString()}
                    </button>
                  ))}
                </div>
              </div>

              <button
                onClick={handleExport}
                disabled={loading}
                style={{
                  width: "100%",
                  background: loading ? "#B0B0B0" : "#1F4E79",
                  color: "white",
                  border: "none",
                  borderRadius: "10px",
                  padding: "16px",
                  fontSize: "16px",
                  fontWeight: "700",
                  cursor: loading ? "not-allowed" : "pointer",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: "10px",
                }}
              >
                {loading ? "Lauft..." : "Export starten"}
              </button>

              {progress && (
                <div style={{
                  marginTop: "14px",
                  padding: "10px 14px",
                  background: "#EBF3FB",
                  borderRadius: "8px",
                  fontSize: "13px",
                  color: "#1F4E79",
                  textAlign: "center",
                }}>
                  {progress}
                </div>
              )}

              {error && (
                <div style={{
                  marginTop: "14px",
                  padding: "10px 14px",
                  background: "#FFF5F5",
                  border: "1px solid #FEB2B2",
                  borderRadius: "8px",
                  fontSize: "13px",
                  color: "#C53030",
                }}>
                  {error}
                </div>
              )}

              {result && (
                <div style={{
                  marginTop: "14px",
                  padding: "14px",
                  background: "#F0FFF4",
                  border: "1px solid #C6F6D5",
                  borderRadius: "8px",
                  textAlign: "center",
                }}>
                  <div style={{ fontWeight: "700", color: "#276749", fontSize: "15px" }}>
                    Excel wurde heruntergeladen!
                  </div>
                  <div style={{ color: "#48BB78", fontSize: "13px", marginTop: "4px" }}>
                    {result.contacts} Kontakte aus {parseInt(result.mails).toLocaleString()} Mails
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        <div style={{
          padding: "12px 32px",
          borderTop: "1px solid #F0F0F0",
          textAlign: "center",
          fontSize: "11px",
          color: "#AAA",
        }}>
          Nur Lesezugriff auf Mails - Keine Daten werden gespeichert
        </div>
      </div>

      <style>{`
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
        * { box-sizing: border-box; }
      `}</style>
    </main>
  )
}
