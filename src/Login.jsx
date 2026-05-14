import { useState } from "react";
import { Lock, User as UserIcon, Eye, EyeOff, ShieldCheck } from "lucide-react";
import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { api, setStoredAuth } from "@/lib/api";
import { cn } from "@/lib/utils";

export default function Login({ onSuccess }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [showPwd, setShowPwd] = useState(false);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    if (!username.trim() || !password) return;
    setBusy(true);
    setError("");
    try {
      const result = await api.login(username.trim(), password);
      setStoredAuth(result);
      onSuccess?.(result);
    } catch (err) {
      setError(err.message || "Login failed");
      setPassword("");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div
      className="flex min-h-screen items-center justify-center px-4 py-10"
      style={{
        background:
          "radial-gradient(80% 60% at 50% -10%, rgba(167,139,250,0.18), transparent 70%)," +
          "radial-gradient(60% 50% at 100% 100%, rgba(124,58,237,0.10), transparent 70%)," +
          "linear-gradient(180deg, #faf9fd 0%, #f0eef7 100%)",
      }}
    >
      <div className="w-full max-w-sm">
        {/* Brand */}
        <div className="mb-7 flex flex-col items-center">
          <div
            className="relative mb-3 flex h-14 w-14 items-center justify-center overflow-hidden rounded-2xl text-white font-extrabold"
            style={{
              background: "linear-gradient(135deg,#a78bfa 0%,#8b5cf6 45%,#6d28d9 100%)",
              boxShadow:
                "0 10px 30px rgba(124,58,237,.45), inset 0 1px 0 rgba(255,255,255,.35), inset 0 -1px 0 rgba(0,0,0,.2)",
              fontSize: 22,
              letterSpacing: "-0.5px",
            }}
          >
            <span className="relative z-10">S</span>
            <div
              className="pointer-events-none absolute inset-0"
              style={{
                background:
                  "radial-gradient(60% 50% at 30% 20%, rgba(255,255,255,.4), transparent 70%)",
              }}
            />
          </div>
          <h1 className="text-[22px] font-bold tracking-tight text-foreground">SettleOps</h1>
          <p className="mt-1 text-xs text-muted-foreground">Sign in to access the dashboard</p>
        </div>

        <Card className="border-border/70 shadow-lg">
          <CardContent className="space-y-4 px-6 py-6">
            <form onSubmit={submit} className="space-y-4">
              <div className="space-y-1.5">
                <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                  Username
                </Label>
                <div className="relative">
                  <UserIcon className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground/70" />
                  <Input
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    placeholder="username"
                    autoComplete="username"
                    autoFocus
                    className="h-11 pl-9"
                    disabled={busy}
                  />
                </div>
              </div>

              <div className="space-y-1.5">
                <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                  Password
                </Label>
                <div className="relative">
                  <Lock className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground/70" />
                  <Input
                    type={showPwd ? "text" : "password"}
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    placeholder="••••••••"
                    autoComplete="current-password"
                    className="h-11 pl-9 pr-10"
                    disabled={busy}
                  />
                  <button
                    type="button"
                    onClick={() => setShowPwd((s) => !s)}
                    tabIndex={-1}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 rounded p-1 text-muted-foreground/70 hover:text-foreground"
                    aria-label={showPwd ? "Hide password" : "Show password"}
                  >
                    {showPwd ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                  </button>
                </div>
              </div>

              {error && (
                <div className="rounded-md border border-red-200 bg-red-50 px-3 py-2 text-xs font-medium text-red-800">
                  {error}
                </div>
              )}

              <Button
                type="submit"
                disabled={busy || !username.trim() || !password}
                className="h-11 w-full text-sm font-semibold tracking-wide"
              >
                {busy ? "Signing in…" : "Sign in"}
              </Button>
            </form>

            <div className="flex items-center gap-1.5 border-t border-border/60 pt-3 text-[11px] text-muted-foreground">
              <ShieldCheck className="h-3 w-3" />
              Single-user access · JWT-secured
            </div>
          </CardContent>
        </Card>

        <p className="mt-6 text-center text-[10.5px] text-muted-foreground">
          SettleOps · Reconcile · Pro
        </p>
      </div>
    </div>
  );
}
