import ShipmentApp from "@/components/shipment-app"

export default function Page() {
  return (
    <main className="min-h-screen bg-background">
      <header className="border-b border-border px-6 py-3 flex items-center gap-3">
        <span className="font-semibold text-sm text-foreground">Shipment Control</span>
        <span className="text-xs text-muted-foreground">Amazon FBA 출하 관리</span>
      </header>
      <div className="px-6">
        <ShipmentApp />
      </div>
    </main>
  )
}
