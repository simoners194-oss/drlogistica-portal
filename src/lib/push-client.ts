// DR Portal — attivazione notifiche push lato client.
// Registra il service worker, chiede il permesso, crea la subscription con la
// chiave VAPID del server e la salva su SharePoint via server function.

import { spGetVapidPublicKey, spSavePushSubscription } from "./sharepoint.functions";

export type PushSupport = "ok" | "insecure" | "unsupported" | "ios-not-installed";

// Le notifiche push richiedono HTTPS, i permessi Notification e (su iOS) che
// la PWA sia stata installata sulla schermata Home.
export function checkPushSupport(): PushSupport {
  if (typeof window === "undefined") return "unsupported";
  if (!window.isSecureContext) return "insecure";
  if (
    !("serviceWorker" in navigator) ||
    !("PushManager" in window) ||
    !("Notification" in window)
  ) {
    // iOS Safari fuori-PWA non espone PushManager: suggerisci l'installazione.
    const isIos = /iphone|ipad|ipod/i.test(navigator.userAgent);
    const standalone =
      window.matchMedia("(display-mode: standalone)").matches ||
      (navigator as unknown as { standalone?: boolean }).standalone === true;
    return isIos && !standalone ? "ios-not-installed" : "unsupported";
  }
  return "ok";
}

export function pushPermission(): NotificationPermission | null {
  if (typeof window === "undefined" || !("Notification" in window)) return null;
  return Notification.permission;
}

function b64urlToUint8(b64: string): Uint8Array {
  const pad = b64.length % 4 === 0 ? "" : "=".repeat(4 - (b64.length % 4));
  const s = atob(b64.replace(/-/g, "+").replace(/_/g, "/") + pad);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}

// Attiva le notifiche su questo dispositivo. Ritorna un messaggio d'errore
// user-friendly, oppure null se tutto ok.
export async function enablePushNotifications(): Promise<string | null> {
  const support = checkPushSupport();
  if (support === "insecure") return "Le notifiche richiedono una connessione sicura (HTTPS).";
  if (support === "ios-not-installed")
    return 'Su iPhone/iPad installa prima l\'app: tasto Condividi → "Aggiungi alla schermata Home", poi riprova da lì.';
  if (support === "unsupported") return "Questo browser non supporta le notifiche push.";

  const permission = await Notification.requestPermission();
  if (permission !== "granted")
    return "Permesso negato: abilita le notifiche per questo sito nelle impostazioni del browser.";

  const reg = await navigator.serviceWorker.register("/sw.js");
  await navigator.serviceWorker.ready;

  const { publicKey } = await spGetVapidPublicKey();
  const existing = await reg.pushManager.getSubscription();
  const sub =
    existing ??
    (await reg.pushManager.subscribe({
      userVisibleOnly: true,
      applicationServerKey: b64urlToUint8(publicKey).buffer as ArrayBuffer,
    }));

  const json = sub.toJSON();
  if (!json.endpoint || !json.keys?.p256dh || !json.keys?.auth)
    return "Subscription non valida: riprova.";
  await spSavePushSubscription({
    data: { endpoint: json.endpoint, p256dh: json.keys.p256dh, auth: json.keys.auth },
  });
  return null;
}
