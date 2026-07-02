import logoAsset from "@/assets/dr-logistica.png.asset.json";

export function Logo({ size = 36, variant = "dark" }: { size?: number; variant?: "dark" | "light" }) {
  return (
    <div className="flex items-center gap-2.5">
      <img src={logoAsset.url} alt="DR Logistica" height={size} style={{ height: size, width: "auto" }} />
      <div className={`leading-tight border-l pl-2.5 ${variant === "light" ? "border-white/30" : "border-border"}`}>
        <div className={`text-[13px] font-semibold tracking-wide uppercase ${variant === "light" ? "text-white" : "text-foreground"}`}>DR Portal</div>
        <div className={`text-[10px] -mt-0.5 ${variant === "light" ? "text-white/70" : "text-muted-foreground"}`}>Presenze</div>
      </div>
    </div>
  );
}