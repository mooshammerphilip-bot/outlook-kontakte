"use client"
import { SessionProvider } from "next-auth/react"

export default function RootLayout({ children }) {
  return (
    <html lang="de">
      <head>
        <title>Outlook Kontakte Exporter</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
      </head>
      <body style={{ margin: 0, padding: 0 }}>
        <SessionProvider>{children}</SessionProvider>
      </body>
    </html>
  )
}
