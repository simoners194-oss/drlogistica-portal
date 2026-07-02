export function Logo({ size = 36 }: { size?: number }) {
  return (
    <div className="flex items-center gap-2.5">
      <img src="/favicon.png" alt="DR Portal" width={size} height={size} className="rounded-lg shadow-sm" />
      <div className="leading-tight">
        <div className="text-[15px] font-semibold tracking-tight text-foreground">DR Portal</div>
        <div className="text-[11px] text-muted-foreground -mt-0.5">DR Logistica</div>
      </div>
    </div>
  );
}