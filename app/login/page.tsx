"use client";

import { useEffect, useState } from "react";

type Mode = "loading" | "login" | "setup";

export default function LoginPage() {
  const [mode, setMode] = useState<Mode>("loading"),
    [showPassword, setShowPassword] = useState(false),
    [busy, setBusy] = useState(false),
    [error, setError] = useState("");

  useEffect(() => {
    fetch("/api/auth/session", { cache: "no-store" })
      .then(async (response) => {
        const result = await response.json();
        if (!response.ok)
          throw new Error(result.error || "Unable to open login");
        if (result.authenticated) window.location.replace("/");
        else setMode(result.needsSetup ? "setup" : "login");
      })
      .catch((reason) => {
        setError(
          reason instanceof Error ? reason.message : "Unable to open login",
        );
        setMode("login");
      });
  }, []);

  const submit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setBusy(true);
    setError("");
    const form = new FormData(event.currentTarget),
      password = String(form.get("password") || ""),
      confirmation = String(form.get("confirm_password") || "");
    if (mode === "setup" && password !== confirmation) {
      setError("Passwords do not match");
      setBusy(false);
      return;
    }
    try {
      const response = await fetch(
          mode === "setup" ? "/api/auth/setup" : "/api/auth/login",
          {
            method: "POST",
            headers: { "content-type": "application/json" },
            body: JSON.stringify({
              name: String(form.get("name") || ""),
              user_id: String(form.get("user_id") || ""),
              password,
            }),
          },
        ),
        result = await response.json();
      if (!response.ok) throw new Error(result.error || "Unable to sign in");
      window.location.replace("/");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "Unable to sign in");
      setBusy(false);
    }
  };

  return (
    <main className="auth-page">
      <section className="auth-visual" aria-label="BillFlow overview">
        <div className="auth-brand">
          <span>≋</span>
          <b>BillFlow</b>
        </div>
        <div className="auth-copy">
          <p>COMPLETE BUSINESS CONTROL</p>
          <h1>Billing, inventory and accounts—connected.</h1>
          <span>
            Manage invoices, purchases, stock, cash, bank, expenses and live
            profit reports from one secure workspace.
          </span>
          <div className="auth-features">
            <article>
              <i>✓</i>
              <div>
                <b>Private access</b>
                <small>Role-based permissions for every user</small>
              </div>
            </article>
            <article>
              <i>✓</i>
              <div>
                <b>Real-time records</b>
                <small>Every module uses the connected database</small>
              </div>
            </article>
            <article>
              <i>✓</i>
              <div>
                <b>Excel ready</b>
                <small>Fast imports with validation and progress</small>
              </div>
            </article>
          </div>
        </div>
        <small className="auth-footer">
          Secure business workspace · BillFlow
        </small>
      </section>
      <section className="auth-form-side">
        <form className="auth-card" onSubmit={submit}>
          <div className="auth-mobile-brand">
            <span>≋</span>
            <b>BillFlow</b>
          </div>
          {mode === "loading" ? (
            <div className="auth-loading">
              <span className="auth-spinner" />
              <h2>Opening secure login</h2>
              <p>Checking your BillFlow workspace…</p>
            </div>
          ) : (
            <>
              <p className="auth-eyebrow">
                {mode === "setup" ? "FIRST-TIME SETUP" : "SECURE SIGN IN"}
              </p>
              <h2>
                {mode === "setup"
                  ? "Create administrator login"
                  : "Welcome back"}
              </h2>
              <span className="auth-subtitle">
                {mode === "setup"
                  ? "Choose the primary User ID and password for BillFlow."
                  : "Enter the User ID and password given by your administrator."}
              </span>
              {mode === "setup" && (
                <label>
                  Full Name
                  <input
                    name="name"
                    autoComplete="name"
                    placeholder="Administrator name"
                    required
                  />
                </label>
              )}
              <label>
                User ID
                <input
                  name="user_id"
                  autoComplete="username"
                  placeholder="Enter your User ID"
                  minLength={3}
                  maxLength={32}
                  autoCapitalize="none"
                  required
                />
              </label>
              <label>
                Password
                <span className="password-field">
                  <input
                    name="password"
                    type={showPassword ? "text" : "password"}
                    autoComplete={
                      mode === "setup" ? "new-password" : "current-password"
                    }
                    placeholder="Enter your password"
                    minLength={8}
                    required
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword((value) => !value)}
                  >
                    {showPassword ? "Hide" : "Show"}
                  </button>
                </span>
              </label>
              {mode === "setup" && (
                <label>
                  Confirm Password
                  <input
                    name="confirm_password"
                    type={showPassword ? "text" : "password"}
                    autoComplete="new-password"
                    placeholder="Re-enter your password"
                    minLength={8}
                    required
                  />
                </label>
              )}
              {error && (
                <div className="auth-error" role="alert">
                  {error}
                </div>
              )}
              <button className="auth-submit" disabled={busy}>
                {busy
                  ? mode === "setup"
                    ? "Creating account…"
                    : "Signing in…"
                  : mode === "setup"
                    ? "Activate BillFlow"
                    : "Sign in to BillFlow"}
              </button>
              <small className="auth-security">
                Your password is securely hashed and never displayed to
                administrators.
              </small>
            </>
          )}
        </form>
      </section>
    </main>
  );
}
