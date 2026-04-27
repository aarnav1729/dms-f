import React, {
  createContext,
  useContext,
  useState,
  useEffect,
  useCallback,
  type ReactNode,
} from "react";

export interface User {
  email: string;
  empId: string | null;
  empName: string | null;
  department: string | null;
  location: string | null;
  reportingManager: string | null;
  reportingManagerEmail?: string | null;
  reportingManagerName?: string | null;
  isAdmin: boolean;
  canUpload?: boolean;
}

interface AuthContextType {
  user: User | null;
  loading: boolean;
  error: string | null;
  logout: () => void;
  refetchSession: () => Promise<void>;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

const DEV_BYPASS_AUTH = String(import.meta.env.VITE_DEV_BYPASS_AUTH || "").toLowerCase() === "true";
const DEV_BYPASS_ADMIN = String(import.meta.env.VITE_DEV_BYPASS_ADMIN || "").toLowerCase() === "true";
const DEV_BYPASS_EMAIL = String(import.meta.env.VITE_DEV_EMAIL || "aarnav.singh@premierenergies.com");

// ── Redirect loop protection ────────────────────────────────────
// If we were redirected from digi and the session STILL fails, we must
// not bounce back again.  We detect this with a short-lived sessionStorage
// flag that survives the cross-domain round-trip.
const REDIRECT_GUARD_KEY = "dms_auth_redirect_ts";
const REDIRECT_COOLDOWN_MS = 15_000; // 15 seconds

function buildDigiLoginUrl(): string {
  return `https://digi.premierenergies.com/login?redirect=${encodeURIComponent(
    window.location.href
  )}&app=dms`;
}

function shouldRedirectToDigi(): boolean {
  const last = sessionStorage.getItem(REDIRECT_GUARD_KEY);
  if (last) {
    const elapsed = Date.now() - Number(last);
    if (elapsed < REDIRECT_COOLDOWN_MS) {
      // We redirected very recently — don't loop
      return false;
    }
  }
  return true;
}

function markRedirectToDigi(): void {
  sessionStorage.setItem(REDIRECT_GUARD_KEY, String(Date.now()));
}

export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const redirectToLogin = useCallback(() => {
    if (!shouldRedirectToDigi()) {
      // We already tried redirecting recently and came back without a valid
      // session.  Show an error instead of looping forever.
      setError("Session could not be established. Please clear cookies and try again, or visit digi.premierenergies.com directly.");
      setLoading(false);
      return;
    }
    markRedirectToDigi();
    window.location.href = buildDigiLoginUrl();
  }, []);

  const fetchSession = useCallback(async () => {
    if (DEV_BYPASS_AUTH) {
      setUser({
        email: DEV_BYPASS_EMAIL,
        empId: "DEV001",
        empName: "Dev User",
        department: "IT",
        location: "Hyderabad",
        reportingManager: "dev.manager@premierenergies.com",
        reportingManagerEmail: "dev.manager@premierenergies.com",
        reportingManagerName: "Dev Manager",
        isAdmin: DEV_BYPASS_ADMIN,
        canUpload: true,
      });
      setLoading(false);
      setError(null);
      return;
    }

    try {
      setLoading(true);
      setError(null);

      const res = await fetch("/api/session", {
        method: "GET",
        credentials: "include",
      });

      if (res.status === 401) {
        // SSO cookie missing or expired.
        //
        // DMS does NOT have its own /auth/refresh endpoint – tokens are
        // issued exclusively by digi.premierenergies.com.  The only way
        // to get a fresh token is to redirect there; digi will auto-login
        // via its own refresh cookie and bounce back with a new sso cookie.
        redirectToLogin();
        return;
      }

      if (!res.ok) {
        throw new Error(`Session fetch failed with status ${res.status}`);
      }

      const data = await res.json();

      // Session succeeded – clear any redirect guard so future 401s can
      // redirect normally.
      sessionStorage.removeItem(REDIRECT_GUARD_KEY);

      setUser({
        email: data.email,
        empId: data.empId,
        empName: data.empName,
        department: data.department,
        location: data.location,
        reportingManager: data.reportingManagerName || data.reportingManagerEmail || data.reportingManagerId || null,
        reportingManagerEmail: data.reportingManagerEmail || null,
        reportingManagerName: data.reportingManagerName || null,
        isAdmin: Boolean(data.isAdmin),
        canUpload: Boolean(data.canUpload || data.isAdmin),
      });
    } catch (err) {
      console.error("Auth session error:", err);
      setError("Failed to verify your session. Redirecting to login...");
      setTimeout(() => redirectToLogin(), 2000);
    } finally {
      setLoading(false);
    }
  }, [redirectToLogin]);

  useEffect(() => {
    fetchSession();
  }, [fetchSession]);

  const logout = useCallback(() => {
    if (DEV_BYPASS_AUTH) {
      setUser(null);
      setTimeout(() => {
        fetchSession();
      }, 0);
      return;
    }

    fetch("/auth/logout", { method: "POST", credentials: "include" })
      .catch(() => {})
      .finally(() => {
        setUser(null);
        // On logout, always allow redirect (clear guard)
        sessionStorage.removeItem(REDIRECT_GUARD_KEY);
        redirectToLogin();
      });
  }, [redirectToLogin, fetchSession]);

  return (
    <AuthContext.Provider
      value={{ user, loading, error, logout, refetchSession: fetchSession }}
    >
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth(): AuthContextType {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error("useAuth must be used within an AuthProvider");
  }
  return context;
}
