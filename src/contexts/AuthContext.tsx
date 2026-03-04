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

const DIGI_LOGIN_URL = `https://digi.premierenergies.com/login?redirect=${encodeURIComponent(
  typeof window !== "undefined" ? window.location.origin : "https://dms.premierenergies.com"
)}&app=dms`;

export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const redirectToLogin = useCallback(() => {
    window.location.href = DIGI_LOGIN_URL;
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
        // SSO cookie missing or expired - try refresh first
        const refreshRes = await fetch("/auth/refresh", {
          method: "POST",
          credentials: "include",
        });

        if (refreshRes.ok) {
          // Retry session after refresh
          const retryRes = await fetch("/api/session", {
            method: "GET",
            credentials: "include",
          });

          if (retryRes.ok) {
            const data = await retryRes.json();
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
            });
            return;
          }
        }

        // Both session and refresh failed - redirect to DIGI login
        redirectToLogin();
        return;
      }

      if (!res.ok) {
        throw new Error(`Session fetch failed with status ${res.status}`);
      }

      const data = await res.json();
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
        redirectToLogin();
      });
  }, [redirectToLogin]);

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
