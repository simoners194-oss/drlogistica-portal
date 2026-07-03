import logoAsset from "@/assets/dr-logistica.png.asset.json";
import { APP_NAME, APP_TAGLINE } from "@/lib/modules";

export function Logo({
  size = 36,
  variant = "dark",
  subtitle = APP_TAGLINE,
}: {
  size?: number;
  variant?: "dark" | "light";
  subtitle?: string;
}) {
  if (variant === "light") {
    // On dark backgrounds, show only the icon mark + white text (the logo's black wordmark would disappear).
    return (
      <div className="flex items-center gap-2.5">
        <img src="/favicon.png" alt="DR Logistica" width={size} height={size} style={{ width: size, height: size }} />
        <div className="leading-tight border-l border-white/30 pl-2.5">
          <div className="text-[13px] font-semibold tracking-wide uppercase text-white">{APP_NAME}</div>
          <div className="text-[10px] -mt-0.5 text-white/70">{subtitle}</div>
        </div>
      </div>
    );
  }
  return (
    <div className="flex items-center gap-2.5">
      <img src={logoAsset.url} alt="DR Logistica" height={size} style={{ height: size, width: "auto" }} />
    </div>
  );
}