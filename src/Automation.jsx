import { useCallback, useEffect, useState } from "react";
import {
  Bot,
  Mail,
  Send,
  Plug,
  RefreshCw,
  CheckCircle2,
  AlertCircle,
  Clock,
  Download,
  Search,
  ExternalLink,
  Inbox,
  Activity,
  UserPlus,
  Trash2,
  Edit3,
  Save,
  X,
  PauseCircle,
  PlayCircle,
} from "lucide-react";
import {
  Card,
  CardHeader,
  CardTitle,
  CardDescription,
  CardContent,
} from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { cn } from "@/lib/utils";
import { api } from "@/lib/api";

// ---------- helpers ----------

function StatusDot({ ok, label }) {
  return (
    <span className="inline-flex items-center gap-1.5 text-xs font-semibold">
      <span
        className={cn(
          "inline-block h-2 w-2 rounded-full",
          ok ? "bg-emerald-500 shadow-[0_0_6px_rgba(16,185,129,.7)]" : "bg-slate-300"
        )}
      />
      <span className={ok ? "text-emerald-700" : "text-muted-foreground"}>{label}</span>
    </span>
  );
}

function formatDateTime(iso) {
  if (!iso) return "—";
  try {
    return new Date(iso).toLocaleString();
  } catch {
    return iso;
  }
}

// Compact toast — no extra dep
function useToasts() {
  const [items, setItems] = useState([]);
  const push = useCallback((text, tone = "info") => {
    const id = Math.random().toString(36).slice(2);
    setItems((p) => [...p, { id, text, tone }]);
    setTimeout(() => setItems((p) => p.filter((t) => t.id !== id)), 4000);
  }, []);
  const node = (
    <div className="pointer-events-none fixed bottom-4 right-4 z-50 flex flex-col gap-2">
      {items.map((t) => (
        <div
          key={t.id}
          className={cn(
            "pointer-events-auto max-w-sm rounded-lg border px-4 py-2.5 text-sm shadow-lg backdrop-blur",
            t.tone === "error"
              ? "border-red-200 bg-red-50/95 text-red-900"
              : t.tone === "success"
              ? "border-emerald-200 bg-emerald-50/95 text-emerald-900"
              : "border-border bg-popover/95 text-foreground"
          )}
        >
          {t.text}
        </div>
      ))}
    </div>
  );
  return { push, node };
}

// ---------- Backend health banner ----------
// Only render when there is a problem to surface. The header "API" pill already
// indicates connected status — no point showing a fat green banner too.

function BackendBanner({ health, loading }) {
  if (loading || health) return null;
  return (
    <Card className="border-red-300/70 bg-red-50/60 shadow-xs">
      <CardContent className="flex flex-wrap items-center gap-3 px-5 py-3 text-sm">
        <AlertCircle className="h-4 w-4 flex-shrink-0 text-red-600" />
        <span className="text-red-900">
          Backend offline. Start it with{" "}
          <code className="rounded bg-red-100 px-1.5 py-0.5 font-mono text-[11px]">
            uvicorn app.main:app --reload
          </code>{" "}
          from the <code className="rounded bg-red-100 px-1.5 py-0.5 font-mono text-[11px]">backend/</code> folder.
        </span>
      </CardContent>
    </Card>
  );
}

// ---------- Telegram panel ----------

const LAST_SYNC_KEY = "settleops:tg:lastSync";

function TelegramPanel({ pushToast }) {
  const [config, setConfig] = useState(null);
  const [me, setMe] = useState(null);
  const [meErr, setMeErr] = useState(null);
  const [busy, setBusy] = useState(false);
  // Persist lastSync across remounts / probe-driven re-renders
  const [lastSync, setLastSyncRaw] = useState(() => {
    try {
      const cached = localStorage.getItem(LAST_SYNC_KEY);
      return cached ? JSON.parse(cached) : null;
    } catch {
      return null;
    }
  });
  const setLastSync = (value) => {
    setLastSyncRaw(value);
    try {
      if (value) localStorage.setItem(LAST_SYNC_KEY, JSON.stringify(value));
      else localStorage.removeItem(LAST_SYNC_KEY);
    } catch {
      /* quota / private mode — ignore */
    }
  };

  const [tokenDraft, setTokenDraft] = useState("");
  const [chatDraft, setChatDraft] = useState("");
  const [allowDraft, setAllowDraft] = useState("");
  const [intervalDraft, setIntervalDraft] = useState(10);
  const [cronDraft, setCronDraft] = useState("");

  const [adhocMsg, setAdhocMsg] = useState("");
  const [adhocTo, setAdhocTo] = useState("");
  const [reportTo, setReportTo] = useState("");

  const loadConfig = useCallback(async () => {
    try {
      const cfg = await api.tgConfig();
      setConfig(cfg);
      setChatDraft(cfg.report_chat_id || "");
      setAllowDraft((cfg.inbound_chat_ids || []).join(","));
      setIntervalDraft(cfg.sync_interval_minutes ?? 10);
      setCronDraft(cfg.digest_cron || "");
    } catch (e) {
      pushToast(`Couldn't load Telegram config: ${e.message}`, "error");
    }
  }, [pushToast]);

  const loadMe = useCallback(async () => {
    setMeErr(null);
    try {
      const u = await api.tgMe();
      setMe(u);
    } catch (e) {
      setMe(null);
      setMeErr(e.message);
    }
  }, []);

  useEffect(() => {
    loadConfig();
    loadMe();
  }, [loadConfig, loadMe]);

  const saveConfig = async () => {
    setBusy(true);
    try {
      const body = {
        sync_interval_minutes: Number(intervalDraft) || 0,
        digest_cron: cronDraft,
        report_chat_id: chatDraft,
        inbound_chat_ids: allowDraft
          .split(",")
          .map((s) => s.trim())
          .filter(Boolean),
      };
      if (tokenDraft) body.bot_token = tokenDraft;
      await api.updateTgConfig(body);
      setTokenDraft("");
      pushToast("Telegram settings saved", "success");
      await loadConfig();
      await loadMe();
    } catch (e) {
      pushToast(`Save failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const doSync = async () => {
    setBusy(true);
    try {
      const result = await api.tgSync();
      setLastSync(result);
      if (result.run.status === "error") {
        pushToast(result.run.error_message || "Sync failed", "error");
        return;
      }
      pushToast(
        `Sync OK · ${result.run.messages_seen} message(s) · ${result.run.uploads_created} upload(s)`,
        "success"
      );
      // If new files came in, reload so the Dashboard KPIs reflect the ingest.
      if (result.run.uploads_created > 0) {
        setTimeout(() => {
          // Reset to default page state (dashboard) on reload
          window.location.reload();
        }, 1200);
      }
    } catch (e) {
      pushToast(`Sync failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const doResetState = async () => {
    if (!confirm("Reset Telegram offset? The next Pull will re-scan all messages Telegram still has (up to ~24h back).")) return;
    setBusy(true);
    try {
      const r = await api.tgResetState();
      pushToast(
        r.had_state ? "Offset cleared — next Pull will re-scan recent messages" : "Offset was already clean",
        "success"
      );
    } catch (e) {
      pushToast(`Reset failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const doSendReport = async () => {
    setBusy(true);
    try {
      const body = { include_excel: true };
      const target = (reportTo || "").trim();
      if (target) body.chat_id = target;
      await api.tgSendReport(body);
      pushToast(target ? `Report sent to ${target}` : "Report sent to default chat", "success");
    } catch (e) {
      pushToast(`Send failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const doSendMessage = async () => {
    if (!adhocMsg.trim()) return;
    setBusy(true);
    try {
      const body = { text: adhocMsg };
      const target = (adhocTo || "").trim();
      if (target) body.chat_id = target;
      await api.tgSendMessage(body);
      setAdhocMsg("");
      pushToast(target ? `Sent to ${target}` : "Sent to default chat", "success");
    } catch (e) {
      pushToast(`Send failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const useChatAsReport = (cid) => {
    setChatDraft(cid);
    pushToast(`Pasted ${cid} into Report chat id — click Save settings to persist`, "info");
  };

  const botReady = !!me;
  const reportSet = !!config?.report_chat_id;
  const needsSetup = !botReady || !reportSet;

  return (
    <div className="space-y-5">
      {/* Compact status strip — single line, scannable */}
      <Card className="border-border/70 shadow-xs">
        <CardContent className="flex flex-wrap items-center gap-x-6 gap-y-3 px-5 py-3.5 text-sm">
          {/* Bot identity */}
          <div className="flex min-w-0 items-center gap-2.5">
            <div className="flex h-7 w-7 items-center justify-center rounded-md bg-sky-100 text-sky-700">
              <Send className="h-3.5 w-3.5" />
            </div>
            {botReady ? (
              <div className="min-w-0 leading-tight">
                <div className="truncate text-[13px] font-semibold text-foreground">
                  {me.first_name}{" "}
                  <span className="font-normal text-muted-foreground">@{me.username}</span>
                </div>
                <div className="font-mono text-[10px] text-muted-foreground">id {me.id}</div>
              </div>
            ) : (
              <div className="text-xs text-muted-foreground">
                {meErr ? `Bot error: ${meErr.slice(0, 60)}` : "Bot not configured"}
              </div>
            )}
          </div>

          <div className="hidden h-7 w-px bg-border/70 sm:block" />

          {/* Report chat status */}
          <div className="flex items-center gap-1.5 text-xs">
            <StatusDot ok={reportSet} label={reportSet ? "Report chat ready" : "Report chat not set"} />
          </div>

          {/* Schedule */}
          <div className="flex items-center gap-3 text-[11px] text-muted-foreground">
            <span>
              Poll <span className="font-mono font-semibold text-foreground">{config?.sync_interval_minutes ?? "—"}m</span>
            </span>
            <span>
              Digest{" "}
              <span className="font-mono font-semibold text-foreground">
                {config?.digest_cron || "off"}
              </span>
            </span>
          </div>

          <Button size="sm" variant="ghost" onClick={loadMe} disabled={busy} className="ml-auto h-7 text-xs">
            <RefreshCw className={cn("mr-1 h-3 w-3", busy && "animate-spin")} /> Recheck
          </Button>
        </CardContent>
      </Card>

      {/* Setup hint banner — only when something's missing */}
      {needsSetup && (
        <Card className="border-amber-200/70 bg-amber-50/50 shadow-xs">
          <CardContent className="flex flex-wrap items-center gap-3 px-5 py-3 text-sm">
            <AlertCircle className="h-4 w-4 flex-shrink-0 text-amber-600" />
            <div className="flex-1 text-amber-900">
              <span className="font-semibold">Setup needed:</span>{" "}
              {!botReady && "paste a bot token below"}
              {!botReady && !reportSet && " and "}
              {!reportSet && "set a default Report chat id (run Pull updates to discover ids)"}
              .
            </div>
          </CardContent>
        </Card>
      )}

      {/* Quick actions — two cards side by side */}
      <div className="grid gap-4 md:grid-cols-2">
        {/* Inbound */}
        <Card className="border-border/70 shadow-xs">
          <CardHeader className="border-b border-border/60 pb-3">
            <div className="flex items-center gap-2.5">
              <div className="flex h-7 w-7 items-center justify-center rounded-md bg-violet-100 text-violet-700">
                <Inbox className="h-3.5 w-3.5" />
              </div>
              <div className="min-w-0 flex-1">
                <CardTitle className="text-[13.5px] font-semibold tracking-tight">Pull bank files</CardTitle>
                <CardDescription className="text-[11px]">
                  Download new <code>.xlsx</code>/<code>.csv</code> files from your bot.
                </CardDescription>
              </div>
            </div>
          </CardHeader>
          <CardContent className="px-5 py-4">
            <div className="flex flex-wrap gap-2">
              <Button onClick={doSync} disabled={busy}>
                <Inbox className="mr-1 h-4 w-4" />
                {busy ? "Working…" : "Pull updates"}
              </Button>
              <Button
                onClick={doResetState}
                disabled={busy}
                variant="outline"
                size="default"
                title="Forget last_update_id so the next Pull re-scans recent messages (helpful if you missed some)"
              >
                <RefreshCw className="mr-1 h-3.5 w-3.5" />
                Reset offset
              </Button>
            </div>

            {lastSync && lastSync.run.status === "ok" && lastSync.run.messages_seen === 0 && (
              <div className="mt-3 rounded-md border border-amber-200 bg-amber-50/60 px-3 py-2 text-[11px] text-amber-900">
                <div className="font-semibold">No new messages since last sync.</div>
                <div className="mt-0.5 text-amber-800/90">
                  Telegram only delivers each update once. If you expected messages, click{" "}
                  <b>Reset offset</b> then Pull again. Also confirm the bot's privacy mode is{" "}
                  <b>disabled</b> in <code>@BotFather</code> → <code>/setprivacy</code> so it sees
                  group files.
                </div>
              </div>
            )}

            {lastSync && (
              <div className="mt-3 rounded-md border border-border bg-muted/40 px-3 py-2 text-xs">
                <div className="mb-1.5 flex items-center justify-between">
                  <span className="text-muted-foreground">Last run</span>
                  <Badge
                    variant="secondary"
                    className={cn(
                      "text-[10px]",
                      lastSync.run.status === "ok"
                        ? "bg-emerald-100 text-emerald-800 hover:bg-emerald-100"
                        : "bg-red-100 text-red-800 hover:bg-red-100"
                    )}
                  >
                    {lastSync.run.status}
                  </Badge>
                </div>
                <div className="grid grid-cols-3 gap-2 text-[11px]">
                  <div>
                    <div className="text-[9.5px] uppercase tracking-wider text-muted-foreground">Msgs</div>
                    <div className="font-mono font-bold tabular-nums">{lastSync.run.messages_seen}</div>
                  </div>
                  <div>
                    <div className="text-[9.5px] uppercase tracking-wider text-muted-foreground">Attach</div>
                    <div className="font-mono font-bold tabular-nums">{lastSync.run.attachments_saved}</div>
                  </div>
                  <div>
                    <div className="text-[9.5px] uppercase tracking-wider text-muted-foreground">Uploads</div>
                    <div className="font-mono font-bold tabular-nums">{lastSync.run.uploads_created}</div>
                  </div>
                </div>
                {lastSync.run.chat_id && (
                  <div className="mt-2 border-t border-border/70 pt-2">
                    <div className="mb-1 text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
                      Chat ids seen — click to use
                    </div>
                    <div className="flex flex-wrap gap-1.5">
                      {lastSync.run.chat_id.split(",").map((cid) => (
                        <button
                          key={cid}
                          type="button"
                          onClick={() => useChatAsReport(cid.trim())}
                          className="inline-flex items-center gap-1 rounded bg-primary/10 px-2 py-0.5 font-mono text-[10.5px] text-primary hover:bg-primary/20"
                        >
                          {cid.trim()} ↑
                        </button>
                      ))}
                    </div>
                  </div>
                )}
                {lastSync.run.error_message && (
                  <div className="mt-2 border-t border-border/70 pt-2 text-[10.5px] text-red-700">
                    {lastSync.run.error_message}
                  </div>
                )}
              </div>
            )}
          </CardContent>
        </Card>

        {/* Outbound */}
        <Card className="border-border/70 shadow-xs">
          <CardHeader className="border-b border-border/60 pb-3">
            <div className="flex items-center gap-2.5">
              <div className="flex h-7 w-7 items-center justify-center rounded-md bg-emerald-100 text-emerald-700">
                <Send className="h-3.5 w-3.5" />
              </div>
              <div className="min-w-0 flex-1">
                <CardTitle className="text-[13.5px] font-semibold tracking-tight">Send a message</CardTitle>
                <CardDescription className="text-[11px]">
                  Leave recipient empty to use the default Report chat.
                </CardDescription>
              </div>
            </div>
          </CardHeader>
          <CardContent className="space-y-4 px-5 py-4">
            {/* Digest send */}
            <div>
              <Label className="mb-1.5 block text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                Send today's digest
              </Label>
              <div className="flex gap-2">
                <Input
                  value={reportTo}
                  onChange={(e) => setReportTo(e.target.value)}
                  placeholder={config?.report_chat_id ? `→ ${config.report_chat_id}` : "Recipient chat id"}
                  className="h-9 font-mono text-xs"
                />
                <Button
                  onClick={doSendReport}
                  size="sm"
                  disabled={busy || (!reportTo.trim() && !config?.report_chat_id)}
                  className="flex-shrink-0"
                >
                  <Send className="mr-1 h-3.5 w-3.5" /> Send digest
                </Button>
              </div>
            </div>

            <div className="h-px bg-border/60" />

            {/* Ad-hoc */}
            <div>
              <Label className="mb-1.5 block text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                Ad-hoc message
              </Label>
              <div className="space-y-2">
                <Input
                  value={adhocTo}
                  onChange={(e) => setAdhocTo(e.target.value)}
                  placeholder={config?.report_chat_id ? `→ ${config.report_chat_id}` : "Recipient chat id"}
                  className="h-9 font-mono text-xs"
                />
                <div className="flex gap-2">
                  <Input
                    value={adhocMsg}
                    onChange={(e) => setAdhocMsg(e.target.value)}
                    placeholder="Hello from SettleOps…"
                    className="h-9"
                  />
                  <Button
                    onClick={doSendMessage}
                    size="sm"
                    disabled={busy || !adhocMsg.trim() || (!adhocTo.trim() && !config?.report_chat_id)}
                    className="flex-shrink-0"
                  >
                    Send
                  </Button>
                </div>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Settings — slim, single column, clear sections */}
      <Card className="border-border/70 shadow-xs">
        <CardHeader className="flex flex-row items-center justify-between border-b border-border/60 pb-3">
          <div>
            <CardTitle className="text-[13.5px] font-semibold tracking-tight">Bot settings</CardTitle>
            <CardDescription className="text-[11px]">
              Runtime changes. Edit <code>backend/.env</code> for permanence.
            </CardDescription>
          </div>
          <Button onClick={saveConfig} disabled={busy} size="sm">
            {busy ? "Saving…" : "Save"}
          </Button>
        </CardHeader>
        <CardContent className="space-y-5 px-5 py-5">
          {/* Identity section */}
          <div className="space-y-2">
            <Label className="text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
              Bot token
            </Label>
            <Input
              type="password"
              value={tokenDraft}
              onChange={(e) => setTokenDraft(e.target.value)}
              placeholder={config?.has_bot_token ? "•••••••• (already set)" : "Paste token from @BotFather"}
              className="h-10 font-mono"
            />
          </div>

          {/* Chats section */}
          <div className="grid gap-4 md:grid-cols-2">
            <div className="space-y-2">
              <Label className="text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                Report chat id (default recipient)
              </Label>
              <Input
                value={chatDraft}
                onChange={(e) => setChatDraft(e.target.value)}
                placeholder="-1001234567890"
                className="h-10 font-mono"
              />
            </div>
            <div className="space-y-2">
              <Label className="text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                Inbound allowlist
              </Label>
              <Input
                value={allowDraft}
                onChange={(e) => setAllowDraft(e.target.value)}
                placeholder="empty = accept all chats"
                className="h-10 font-mono"
              />
            </div>
          </div>

          {/* Schedule section */}
          <div className="grid gap-4 md:grid-cols-2">
            <div className="space-y-2">
              <Label className="text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                Poll every (minutes)
              </Label>
              <Input
                type="number"
                value={intervalDraft}
                onChange={(e) => setIntervalDraft(e.target.value)}
                min={0}
                className="h-10 font-mono"
              />
            </div>
            <div className="space-y-2">
              <Label className="text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                Daily digest cron
              </Label>
              <Input
                value={cronDraft}
                onChange={(e) => setCronDraft(e.target.value)}
                placeholder="30 18 * * * (optional)"
                className="h-10 font-mono"
              />
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

// ---------- Gmail panel ----------

function GmailPanel({ pushToast }) {
  const [config, setConfig] = useState(null);
  const [queryDraft, setQueryDraft] = useState("");
  const [intervalDraft, setIntervalDraft] = useState(30);
  const [messages, setMessages] = useState(null);
  const [busy, setBusy] = useState(false);
  const [lastSync, setLastSync] = useState(null);

  const loadConfig = useCallback(async () => {
    try {
      const cfg = await api.gmailConfig();
      setConfig(cfg);
      setQueryDraft(cfg.search_query || "");
      setIntervalDraft(cfg.sync_interval_minutes ?? 30);
    } catch (e) {
      pushToast(`Couldn't load Gmail config: ${e.message}`, "error");
    }
  }, [pushToast]);

  useEffect(() => {
    loadConfig();
  }, [loadConfig]);

  const saveConfig = async () => {
    setBusy(true);
    try {
      await api.updateGmailConfig({
        search_query: queryDraft,
        sync_interval_minutes: Number(intervalDraft) || 0,
      });
      pushToast("Gmail settings saved", "success");
      await loadConfig();
    } catch (e) {
      pushToast(`Save failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const connect = async () => {
    setBusy(true);
    try {
      const r = await api.gmailAuthStart();
      window.open(r.auth_url, "_blank", "noopener");
      pushToast("Approve the Google consent screen, then come back here.", "info");
    } catch (e) {
      pushToast(`Couldn't start OAuth: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const doSync = async () => {
    setBusy(true);
    try {
      const result = await api.gmailSync();
      setLastSync(result);
      if (result.run.status === "error") {
        pushToast(result.run.error_message || "Sync failed", "error");
      } else {
        pushToast(
          `Sync OK · ${result.run.messages_seen} message(s) · ${result.run.uploads_created} upload(s)`,
          "success"
        );
      }
    } catch (e) {
      pushToast(`Sync failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const previewMessages = async () => {
    setBusy(true);
    try {
      const r = await api.gmailMessages({ query: queryDraft, limit: 10 });
      setMessages(r);
    } catch (e) {
      setMessages([]);
      pushToast(e.message, "error");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="space-y-5">
      {/* Status */}
      <Card className="border-border/70 shadow-xs">
        <CardHeader className="border-b border-border/60 pb-4">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-rose-100 text-rose-700">
              <Mail className="h-5 w-5" />
            </div>
            <div className="min-w-0 flex-1">
              <CardTitle className="text-base font-semibold tracking-tight">Gmail</CardTitle>
              <CardDescription className="text-xs">
                Pulls bank statement Excel attachments using a Gmail search query.
              </CardDescription>
            </div>
          </div>
        </CardHeader>
        <CardContent className="grid gap-4 px-5 py-5 md:grid-cols-2 lg:grid-cols-4">
          <div>
            <div className="mb-1 text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
              OAuth credentials
            </div>
            <StatusDot
              ok={!!config?.has_credentials}
              label={config?.has_credentials ? "credentials.json present" : "credentials.json missing"}
            />
          </div>
          <div>
            <div className="mb-1 text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
              Authorized
            </div>
            <StatusDot
              ok={!!config?.has_token}
              label={config?.has_token ? "Token saved" : "Not connected"}
            />
          </div>
          <div>
            <div className="mb-1 text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
              Search query
            </div>
            <div className="line-clamp-2 font-mono text-[11px] text-foreground">
              {config?.search_query || "—"}
            </div>
          </div>
          <div>
            <div className="mb-1 text-[10.5px] font-semibold uppercase tracking-wider text-muted-foreground">
              Poll
            </div>
            <div className="text-xs">
              Every <span className="font-mono font-semibold">{config?.sync_interval_minutes ?? "—"}m</span>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Actions */}
      <div className="grid gap-4 md:grid-cols-2">
        <Card className="border-border/70 shadow-xs">
          <CardHeader className="pb-3">
            <CardTitle className="text-sm font-semibold tracking-tight">Connect Gmail</CardTitle>
            <CardDescription className="text-xs">
              Opens the Google consent screen. You'll grant the backend read-only Gmail access.
            </CardDescription>
          </CardHeader>
          <CardContent className="px-5 pb-5">
            <Button onClick={connect} disabled={busy || !config?.has_credentials}>
              <Plug className="mr-1 h-4 w-4" />
              {config?.has_token ? "Re-authorize" : "Connect"}
              <ExternalLink className="ml-1 h-3.5 w-3.5" />
            </Button>
            {!config?.has_credentials && (
              <div className="mt-2 text-[11px] text-amber-700">
                Put your <code>credentials.json</code> in <code>backend/secrets/</code> first.
              </div>
            )}
          </CardContent>
        </Card>

        <Card className="border-border/70 shadow-xs">
          <CardHeader className="pb-3">
            <CardTitle className="text-sm font-semibold tracking-tight">Pull statements now</CardTitle>
            <CardDescription className="text-xs">
              Runs the saved Gmail query and parses any Excel attachments.
            </CardDescription>
          </CardHeader>
          <CardContent className="px-5 pb-5">
            <Button onClick={doSync} disabled={busy || !config?.has_token}>
              <Inbox className="mr-1 h-4 w-4" />
              {busy ? "Working…" : "Pull"}
            </Button>
            {lastSync && (
              <div className="mt-3 rounded-md border border-border bg-muted/40 px-3 py-2 text-xs">
                <div className="flex items-center justify-between">
                  <span>Last run</span>
                  <Badge
                    variant="secondary"
                    className={cn(
                      lastSync.run.status === "ok"
                        ? "bg-emerald-100 text-emerald-800 hover:bg-emerald-100"
                        : "bg-red-100 text-red-800 hover:bg-red-100"
                    )}
                  >
                    {lastSync.run.status}
                  </Badge>
                </div>
                <div className="mt-1 grid grid-cols-3 gap-2 text-[11px]">
                  <span>Msgs: <b>{lastSync.run.messages_seen}</b></span>
                  <span>Attach: <b>{lastSync.run.attachments_saved}</b></span>
                  <span>Uploads: <b>{lastSync.run.uploads_created}</b></span>
                </div>
                {lastSync.run.error_message && (
                  <div className="mt-1 text-[11px] text-red-700">{lastSync.run.error_message}</div>
                )}
              </div>
            )}
          </CardContent>
        </Card>
      </div>

      {/* Settings */}
      <Card className="border-border/70 shadow-xs">
        <CardHeader className="border-b border-border/60 pb-4">
          <CardTitle className="text-sm font-semibold tracking-tight">Settings</CardTitle>
          <CardDescription className="text-xs">
            Runtime overrides. Edit <code>backend/.env</code> for permanence.
          </CardDescription>
        </CardHeader>
        <CardContent className="grid gap-4 px-5 py-5">
          <div className="space-y-2">
            <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
              Gmail search query
            </Label>
            <Input
              value={queryDraft}
              onChange={(e) => setQueryDraft(e.target.value)}
              placeholder="from:bank has:attachment newer_than:2d"
              className="h-10 font-mono"
            />
            <p className="text-[10.5px] text-muted-foreground">
              Uses Gmail's regular search syntax.
            </p>
          </div>
          <div className="space-y-2">
            <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
              Poll every (minutes)
            </Label>
            <Input
              type="number"
              value={intervalDraft}
              onChange={(e) => setIntervalDraft(e.target.value)}
              min={0}
              className="h-10 font-mono"
            />
          </div>
          <div className="flex gap-2">
            <Button onClick={saveConfig} disabled={busy}>
              {busy ? "Saving…" : "Save settings"}
            </Button>
            <Button onClick={previewMessages} variant="outline" disabled={busy || !config?.has_token}>
              <Search className="mr-1 h-3.5 w-3.5" /> Preview matches
            </Button>
          </div>
        </CardContent>
      </Card>

      {/* Message preview */}
      {messages !== null && (
        <Card className="border-border/70 shadow-xs">
          <CardHeader className="pb-3">
            <CardTitle className="text-sm font-semibold tracking-tight">
              Matching messages ({messages.length})
            </CardTitle>
          </CardHeader>
          <CardContent className="px-5 pb-5">
            {messages.length === 0 ? (
              <div className="py-4 text-center text-xs text-muted-foreground">No matches.</div>
            ) : (
              <div className="space-y-2">
                {messages.map((m) => (
                  <div
                    key={m.id}
                    className="rounded-md border border-border bg-muted/30 px-3 py-2 text-xs"
                  >
                    <div className="font-semibold text-foreground">{m.subject || "(no subject)"}</div>
                    <div className="text-muted-foreground">{m.sender}</div>
                    {m.attachment_names?.length > 0 && (
                      <div className="mt-1 flex flex-wrap gap-1">
                        {m.attachment_names.map((n) => (
                          <Badge key={n} variant="secondary" className="bg-blue-100 text-blue-800 hover:bg-blue-100">
                            {n}
                          </Badge>
                        ))}
                      </div>
                    )}
                    {m.snippet && (
                      <div className="mt-1 line-clamp-2 text-[10.5px] text-muted-foreground">{m.snippet}</div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </CardContent>
        </Card>
      )}
    </div>
  );
}

// ---------- Runs panel ----------

function RunsPanel({ pushToast }) {
  const [tgRuns, setTgRuns] = useState([]);
  const [gmRuns, setGmRuns] = useState([]);
  const [busy, setBusy] = useState(false);

  const load = useCallback(async () => {
    setBusy(true);
    try {
      const [tg, gm] = await Promise.allSettled([api.tgRuns(30), api.gmailRuns(30)]);
      setTgRuns(tg.status === "fulfilled" ? tg.value : []);
      setGmRuns(gm.status === "fulfilled" ? gm.value : []);
    } finally {
      setBusy(false);
    }
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  const allRuns = [
    ...tgRuns.map((r) => ({ ...r, channel: "telegram", label: r.kind })),
    ...gmRuns.map((r) => ({
      ...r,
      channel: "gmail",
      label: r.triggered_by === "manual" ? "inbound_manual" : "inbound_scheduler",
    })),
  ].sort((a, b) => new Date(b.started_at) - new Date(a.started_at));

  return (
    <Card className="border-border/70 shadow-xs">
      <CardHeader className="flex flex-row items-center justify-between border-b border-border/60 pb-4">
        <div>
          <CardTitle className="text-sm font-semibold tracking-tight">Run history</CardTitle>
          <CardDescription className="text-xs">
            Every Gmail and Telegram automation run with status and counts.
          </CardDescription>
        </div>
        <Button size="sm" variant="outline" onClick={load} disabled={busy}>
          <RefreshCw className={cn("mr-1 h-3.5 w-3.5", busy && "animate-spin")} /> Refresh
        </Button>
      </CardHeader>
      <CardContent className="px-0 py-0">
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead className="bg-muted/40">
              <tr>
                {["When", "Channel", "Kind", "Status", "Msgs", "Attach", "Uploads", "Error"].map((h) => (
                  <th
                    key={h}
                    className={cn(
                      "whitespace-nowrap px-3 py-2.5 text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground",
                      ["Msgs", "Attach", "Uploads"].includes(h) ? "text-right" : "text-left"
                    )}
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {allRuns.length === 0 && (
                <tr>
                  <td colSpan={8} className="px-4 py-6 text-center text-muted-foreground">
                    No runs yet.
                  </td>
                </tr>
              )}
              {allRuns.map((r) => (
                <tr
                  key={`${r.channel}-${r.id}`}
                  className="border-t border-border/60 transition-colors hover:bg-muted/30"
                >
                  <td className="whitespace-nowrap px-3 py-2 text-foreground">
                    <span className="inline-flex items-center gap-1.5">
                      <Clock className="h-3 w-3 text-muted-foreground" />
                      {formatDateTime(r.started_at)}
                    </span>
                  </td>
                  <td className="px-3 py-2">
                    <Badge
                      variant="secondary"
                      className={cn(
                        "text-[10px]",
                        r.channel === "telegram"
                          ? "bg-sky-100 text-sky-800 hover:bg-sky-100"
                          : "bg-rose-100 text-rose-800 hover:bg-rose-100"
                      )}
                    >
                      {r.channel}
                    </Badge>
                  </td>
                  <td className="whitespace-nowrap px-3 py-2 font-mono text-[10.5px] text-muted-foreground">
                    {r.label}
                  </td>
                  <td className="px-3 py-2">
                    <Badge
                      variant="secondary"
                      className={cn(
                        "text-[10px]",
                        r.status === "ok"
                          ? "bg-emerald-100 text-emerald-800 hover:bg-emerald-100"
                          : r.status === "error"
                          ? "bg-red-100 text-red-800 hover:bg-red-100"
                          : "bg-amber-100 text-amber-800 hover:bg-amber-100"
                      )}
                    >
                      {r.status}
                    </Badge>
                  </td>
                  <td className="px-3 py-2 text-right font-mono tabular-nums">{r.messages_seen ?? 0}</td>
                  <td className="px-3 py-2 text-right font-mono tabular-nums">{r.attachments_saved ?? 0}</td>
                  <td className="px-3 py-2 text-right font-mono tabular-nums">{r.uploads_created ?? 0}</td>
                  <td className="max-w-[260px] truncate px-3 py-2 text-[10.5px] text-red-700" title={r.error_message || ""}>
                    {r.error_message || ""}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </CardContent>
    </Card>
  );
}

// ---------- Recipients panel ----------

const emptyRecipient = {
  name: "",
  chat_id: "",
  description: "",
  is_active: true,
  digest_cron: "",
  include_excel: true,
};

function RecipientRow({ recipient, onChange, onDelete, onSave, onSendNow, busy, pushToast }) {
  const [draft, setDraft] = useState(recipient);
  const [editing, setEditing] = useState(false);

  useEffect(() => {
    setDraft(recipient);
  }, [recipient]);

  const update = (patch) => setDraft((d) => ({ ...d, ...patch }));

  const handleSave = async () => {
    const patch = {};
    if (draft.name !== recipient.name) patch.name = draft.name;
    if (draft.chat_id !== recipient.chat_id) patch.chat_id = draft.chat_id;
    if ((draft.description || "") !== (recipient.description || "")) patch.description = draft.description;
    if (draft.is_active !== recipient.is_active) patch.is_active = draft.is_active;
    if (draft.digest_cron !== recipient.digest_cron) patch.digest_cron = draft.digest_cron;
    if (draft.include_excel !== recipient.include_excel) patch.include_excel = draft.include_excel;
    if (Object.keys(patch).length === 0) {
      setEditing(false);
      return;
    }
    await onSave(recipient.id, patch);
    setEditing(false);
  };

  return (
    <div className={cn(
      "rounded-lg border bg-card px-4 py-3 transition-colors",
      recipient.is_active ? "border-border" : "border-border/60 bg-muted/30"
    )}>
      <div className="flex flex-wrap items-start gap-3">
        <div className="flex h-9 w-9 flex-shrink-0 items-center justify-center rounded-lg bg-primary/10 text-primary font-bold uppercase">
          {(recipient.name || "?").slice(0, 1)}
        </div>
        {editing ? (
          <div className="flex-1 min-w-[200px] space-y-2">
            <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
              <Input
                value={draft.name}
                onChange={(e) => update({ name: e.target.value })}
                placeholder="Name"
                className="h-9"
              />
              <Input
                value={draft.chat_id}
                onChange={(e) => update({ chat_id: e.target.value })}
                placeholder="chat_id"
                className="h-9 font-mono"
              />
            </div>
            <Input
              value={draft.description || ""}
              onChange={(e) => update({ description: e.target.value })}
              placeholder="Notes (optional)"
              className="h-9"
            />
            <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
              <Input
                value={draft.digest_cron}
                onChange={(e) => update({ digest_cron: e.target.value })}
                placeholder="30 18 * * * (= 18:30 daily; empty = off)"
                className="h-9 font-mono"
              />
              <div className="flex items-center gap-3 text-xs">
                <label className="inline-flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    checked={draft.is_active}
                    onChange={(e) => update({ is_active: e.target.checked })}
                  />
                  Active
                </label>
                <label className="inline-flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    checked={draft.include_excel}
                    onChange={(e) => update({ include_excel: e.target.checked })}
                  />
                  Include Excel
                </label>
              </div>
            </div>
          </div>
        ) : (
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <div className="text-sm font-semibold text-foreground">{recipient.name}</div>
              <Badge
                variant="secondary"
                className={cn(
                  "text-[9.5px]",
                  recipient.is_active
                    ? "bg-emerald-100 text-emerald-800 hover:bg-emerald-100"
                    : "bg-slate-200 text-slate-700 hover:bg-slate-200"
                )}
              >
                {recipient.is_active ? "Active" : "Paused"}
              </Badge>
              {recipient.digest_cron && (
                <Badge variant="secondary" className="bg-primary/10 font-mono text-[9.5px] text-primary hover:bg-primary/10">
                  <Clock className="mr-1 h-2.5 w-2.5" /> {recipient.digest_cron}
                </Badge>
              )}
            </div>
            <div className="mt-0.5 flex flex-wrap items-center gap-2 text-[11px] text-muted-foreground">
              <span className="font-mono">{recipient.chat_id}</span>
              {recipient.description && <span>· {recipient.description}</span>}
              {!recipient.include_excel && <span>· text only</span>}
            </div>
          </div>
        )}

        <div className="flex flex-shrink-0 flex-wrap gap-1.5">
          {editing ? (
            <>
              <Button size="sm" onClick={handleSave} disabled={busy} className="h-8">
                <Save className="mr-1 h-3.5 w-3.5" /> Save
              </Button>
              <Button
                size="sm"
                variant="outline"
                onClick={() => { setDraft(recipient); setEditing(false); }}
                className="h-8"
                disabled={busy}
              >
                Cancel
              </Button>
            </>
          ) : (
            <>
              <Button
                size="sm"
                variant="outline"
                onClick={() => onSendNow(recipient.id)}
                disabled={busy || !recipient.is_active}
                className="h-8"
                title="Send digest + Excel now"
              >
                <Send className="mr-1 h-3.5 w-3.5" /> Send
              </Button>
              <Button
                size="sm"
                variant="outline"
                onClick={() => onChange(recipient.id, { is_active: !recipient.is_active })}
                disabled={busy}
                className="h-8"
                title={recipient.is_active ? "Pause" : "Resume"}
              >
                {recipient.is_active ? <PauseCircle className="h-3.5 w-3.5" /> : <PlayCircle className="h-3.5 w-3.5" />}
              </Button>
              <Button
                size="sm"
                variant="ghost"
                onClick={() => setEditing(true)}
                disabled={busy}
                className="h-8"
              >
                <Edit3 className="h-3.5 w-3.5" />
              </Button>
              <Button
                size="sm"
                variant="ghost"
                onClick={() => onDelete(recipient.id, recipient.name)}
                disabled={busy}
                className="h-8 text-destructive hover:bg-destructive/10 hover:text-destructive"
              >
                <Trash2 className="h-3.5 w-3.5" />
              </Button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function RecipientsPanel({ pushToast }) {
  const [list, setList] = useState([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [showAdd, setShowAdd] = useState(false);
  const [draft, setDraft] = useState(emptyRecipient);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const rows = await api.tgRecipients();
      setList(rows);
    } catch (e) {
      pushToast(`Couldn't load recipients: ${e.message}`, "error");
    } finally {
      setLoading(false);
    }
  }, [pushToast]);

  useEffect(() => {
    load();
  }, [load]);

  const addRecipient = async () => {
    if (!draft.name.trim() || !draft.chat_id.trim()) {
      pushToast("Name and chat_id are required", "error");
      return;
    }
    setBusy(true);
    try {
      await api.tgAddRecipient(draft);
      pushToast(`Added ${draft.name}`, "success");
      setDraft(emptyRecipient);
      setShowAdd(false);
      await load();
    } catch (e) {
      pushToast(`Add failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const patchRecipient = async (id, patch) => {
    setBusy(true);
    try {
      await api.tgUpdateRecipient(id, patch);
      pushToast("Saved", "success");
      await load();
    } catch (e) {
      pushToast(`Update failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const deleteRecipient = async (id, name) => {
    if (!confirm(`Remove "${name}"?`)) return;
    setBusy(true);
    try {
      await api.tgDeleteRecipient(id);
      pushToast(`Removed ${name}`, "success");
      await load();
    } catch (e) {
      pushToast(`Delete failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const sendNow = async (id) => {
    setBusy(true);
    try {
      await api.tgSendReportTo(id);
      pushToast("Report sent", "success");
    } catch (e) {
      pushToast(`Send failed: ${e.message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="space-y-5">
      <Card className="border-border/70 shadow-xs">
        <CardHeader className="flex flex-row items-start justify-between gap-3 border-b border-border/60 pb-4">
          <div>
            <CardTitle className="text-sm font-semibold tracking-tight">Recipient contact book</CardTitle>
            <CardDescription className="text-xs">
              Saved Telegram users / groups. Each can have its own digest cron schedule.
            </CardDescription>
          </div>
          <Button size="sm" onClick={() => setShowAdd((s) => !s)} disabled={busy}>
            <UserPlus className="mr-1 h-3.5 w-3.5" />
            {showAdd ? "Close" : "Add recipient"}
          </Button>
        </CardHeader>
        {showAdd && (
          <CardContent className="border-b border-border/60 px-5 py-5">
            <div className="grid gap-3 sm:grid-cols-2">
              <div className="space-y-1.5">
                <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                  Name
                </Label>
                <Input
                  value={draft.name}
                  onChange={(e) => setDraft({ ...draft, name: e.target.value })}
                  placeholder="e.g. Aryan, Ops Team"
                  className="h-9"
                />
              </div>
              <div className="space-y-1.5">
                <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                  Chat id
                </Label>
                <Input
                  value={draft.chat_id}
                  onChange={(e) => setDraft({ ...draft, chat_id: e.target.value })}
                  placeholder="5847291046 (user) or -1001234567890 (group)"
                  className="h-9 font-mono"
                />
              </div>
              <div className="space-y-1.5 sm:col-span-2">
                <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                  Notes (optional)
                </Label>
                <Input
                  value={draft.description}
                  onChange={(e) => setDraft({ ...draft, description: e.target.value })}
                  placeholder="Free-form description"
                  className="h-9"
                />
              </div>
              <div className="space-y-1.5">
                <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                  Daily digest cron (optional)
                </Label>
                <Input
                  value={draft.digest_cron}
                  onChange={(e) => setDraft({ ...draft, digest_cron: e.target.value })}
                  placeholder="30 18 * * *"
                  className="h-9 font-mono"
                />
                <p className="text-[10px] text-muted-foreground">
                  Standard 5-field cron. Leave empty for no schedule.
                </p>
              </div>
              <div className="flex flex-col gap-2 text-xs">
                <label className="inline-flex items-center gap-2 mt-2">
                  <input
                    type="checkbox"
                    checked={draft.is_active}
                    onChange={(e) => setDraft({ ...draft, is_active: e.target.checked })}
                  />
                  Active (start scheduling immediately)
                </label>
                <label className="inline-flex items-center gap-2">
                  <input
                    type="checkbox"
                    checked={draft.include_excel}
                    onChange={(e) => setDraft({ ...draft, include_excel: e.target.checked })}
                  />
                  Attach Excel report
                </label>
              </div>
            </div>
            <div className="mt-4 flex gap-2">
              <Button onClick={addRecipient} disabled={busy}>
                <Save className="mr-1 h-4 w-4" /> Save recipient
              </Button>
              <Button variant="outline" onClick={() => { setDraft(emptyRecipient); setShowAdd(false); }} disabled={busy}>
                Cancel
              </Button>
            </div>
          </CardContent>
        )}
        <CardContent className="px-5 py-5">
          {loading ? (
            <div className="py-8 text-center text-sm text-muted-foreground">Loading…</div>
          ) : list.length === 0 ? (
            <div className="rounded-lg border border-dashed border-border/80 bg-background/40 px-6 py-10 text-center">
              <UserPlus className="mx-auto mb-2 h-8 w-8 text-muted-foreground/60" />
              <div className="text-sm font-semibold text-foreground">No recipients yet</div>
              <div className="mb-4 mt-1 text-xs text-muted-foreground">
                Add named contacts so you can send to them by name and schedule per-user digests.
              </div>
              <Button onClick={() => setShowAdd(true)} size="sm">
                <UserPlus className="mr-1 h-4 w-4" /> Add your first recipient
              </Button>
            </div>
          ) : (
            <div className="space-y-2">
              {list.map((r) => (
                <RecipientRow
                  key={r.id}
                  recipient={r}
                  onChange={patchRecipient}
                  onDelete={deleteRecipient}
                  onSave={patchRecipient}
                  onSendNow={sendNow}
                  busy={busy}
                  pushToast={pushToast}
                />
              ))}
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}

// ---------- Main page ----------

export default function Automation() {
  const [tab, setTab] = useState("telegram");
  const [health, setHealth] = useState(null);
  const [loading, setLoading] = useState(true);
  const toasts = useToasts();

  const checkHealth = useCallback(async () => {
    setLoading(true);
    try {
      const h = await api.health();
      setHealth(h);
    } catch {
      setHealth(null);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    checkHealth();
    const id = setInterval(checkHealth, 30_000);
    return () => clearInterval(id);
  }, [checkHealth]);

  return (
    <div className="fade-up space-y-5">
      <BackendBanner health={health} loading={loading} />

      {health && (
        <>
          <div className="-mx-1 max-w-full overflow-x-auto px-1">
            <Tabs value={tab} onValueChange={setTab}>
              <TabsList>
                <TabsTrigger value="telegram">
                  <Send className="mr-1.5 h-3.5 w-3.5" /> Telegram
                </TabsTrigger>
                <TabsTrigger value="recipients">
                  <UserPlus className="mr-1.5 h-3.5 w-3.5" /> Recipients
                </TabsTrigger>
                <TabsTrigger value="gmail">
                  <Mail className="mr-1.5 h-3.5 w-3.5" /> Gmail
                </TabsTrigger>
                <TabsTrigger value="runs">
                  <Activity className="mr-1.5 h-3.5 w-3.5" /> Run history
                </TabsTrigger>
              </TabsList>
            </Tabs>
          </div>

          {tab === "telegram" && <TelegramPanel pushToast={toasts.push} />}
          {tab === "recipients" && <RecipientsPanel pushToast={toasts.push} />}
          {tab === "gmail" && <GmailPanel pushToast={toasts.push} />}
          {tab === "runs" && <RunsPanel pushToast={toasts.push} />}
        </>
      )}

      {toasts.node}
    </div>
  );
}
