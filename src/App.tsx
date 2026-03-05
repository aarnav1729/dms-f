import React, {
  useEffect,
  useMemo,
  useState,
  useCallback,
  useRef,
} from "react";
import {
  BrowserRouter,
  Link,
  NavLink,
  Route,
  Routes,
  useLocation,
  useNavigate,
  useParams,
} from "react-router-dom";
import { AnimatePresence, motion } from "framer-motion";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
} from "recharts";
import {
  Activity,
  BookOpen,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  Clock3,
  Download,
  Eye,
  FileArchive,
  FileCheck2,
  FileText,
  FileUp,
  FolderOpen,
  Grid3X3,
  Hash,
  LayoutDashboard,
  List,
  LogOut,
  Menu,
  MoreHorizontal,
  PanelLeftClose,
  PanelLeftOpen,
  Search,
  Settings,
  Share2,
  Shield,
  ShieldCheck,
  Sparkles,
  Upload,
  UserPlus2,
  Users,
  X,
  Zap,
  ExternalLink,
  Copy,
  AlertTriangle,
  Printer,
  Lock,
  Trash2,
  PenLine,
  Plus,
} from "lucide-react";
import { AuthProvider, useAuth } from "@/contexts/AuthContext";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import { Badge } from "@/components/ui/badge";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

/* ──────────────────────────────────────────────────────────
   Types
   ────────────────────────────────────────────────────────── */
type Doc = {
  Id: number;
  Title: string;
  Description: string;
  FileName: string;
  CurrentVersion: number;
  CurrentVersionLabel?: string;
  IsControlled: boolean;
  Status: string;
  ShareScope: string;
  CreatorEmail: string;
  Department: string;
  Location: string;
  ApprovalStatus: string;
  HodSkipped?: boolean;
  CreatedAt: string;
  UpdatedAt: string;
  ValidFrom?: string;
  ValidTo?: string;
};
type Approval = {
  Id: number;
  DocId: number;
  Stage: string;
  Status: string;
  Title: string;
  CurrentVersion: number;
  CreatorEmail: string;
  Department: string;
  Location: string;
  HodSkipped?: boolean;
};
type Employee = {
  EmpID: string;
  EmpName: string;
  EmpEmail: string;
  Department?: string;
  Dept?: string;
  Location?: string;
  EmpLocation?: string;
  ReportingManagerID?: string;
  ManagerID?: string;
};

/* ──────────────────────────────────────────────────────────
   API Helper
   ────────────────────────────────────────────────────────── */
async function api<T>(url: string, init?: RequestInit): Promise<T> {
  const r = await fetch(url, { credentials: "include", ...init });
  if (!r.ok) throw new Error(await r.text().catch(() => `HTTP ${r.status}`));
  return r.json() as Promise<T>;
}

const BASE =
  typeof window !== "undefined"
    ? window.location.origin
    : "https://dms.premierenergies.com";

/* ──────────────────────────────────────────────────────────
   Toast System
   ────────────────────────────────────────────────────────── */
type AppToast = {
  id: number;
  title: string;
  description?: string;
  variant?: "success" | "error" | "info";
};
const ToastCtx = React.createContext<
  ((t: Omit<AppToast, "id">) => void) | undefined
>(undefined);
function useAppToast() {
  const c = React.useContext(ToastCtx);
  if (!c) throw new Error("wrap in ToastProvider");
  return c;
}

function ToastProvider({ children }: { children: React.ReactNode }) {
  const [toasts, setToasts] = useState<AppToast[]>([]);
  const push = (t: Omit<AppToast, "id">) => {
    const id = Date.now() + Math.random();
    setToasts((p) => [...p, { id, ...t }]);
    setTimeout(() => setToasts((p) => p.filter((x) => x.id !== id)), 4000);
  };
  return (
    <ToastCtx.Provider value={push}>
      {children}
      <div className="fixed top-4 right-4 z-[200] space-y-2 w-[min(92vw,380px)]">
        <AnimatePresence>
          {toasts.map((t) => (
            <motion.div
              key={t.id}
              initial={{ opacity: 0, y: -10, scale: 0.98 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: -6, scale: 0.98 }}
              className={`rounded-xl border p-3.5 shadow-xl backdrop-blur-xl text-sm ${
                t.variant === "error"
                  ? "bg-red-50/95 border-red-200 text-red-900"
                  : t.variant === "success"
                  ? "bg-emerald-50/95 border-emerald-200 text-emerald-900"
                  : "bg-white/95 border-slate-200 text-slate-900"
              }`}
            >
              <div className="flex items-start justify-between gap-2">
                <div>
                  <div className="font-semibold">{t.title}</div>
                  {t.description ? (
                    <div className="text-xs opacity-75 mt-0.5">
                      {t.description}
                    </div>
                  ) : null}
                </div>
                <button
                  className="opacity-50 hover:opacity-100 text-xs mt-0.5"
                  onClick={() =>
                    setToasts((p) => p.filter((x) => x.id !== t.id))
                  }
                >
                  ✕
                </button>
              </div>
            </motion.div>
          ))}
        </AnimatePresence>
      </div>
    </ToastCtx.Provider>
  );
}

/* ──────────────────────────────────────────────────────────
   Mouse Glow
   ────────────────────────────────────────────────────────── */
function MouseGlow() {
  useEffect(() => {
    const h = (e: MouseEvent) => {
      document.documentElement.style.setProperty(
        "--spotlight-x",
        `${(e.clientX / window.innerWidth) * 100}%`
      );
      document.documentElement.style.setProperty(
        "--spotlight-y",
        `${(e.clientY / window.innerHeight) * 100}%`
      );
    };
    window.addEventListener("mousemove", h, { passive: true });
    return () => window.removeEventListener("mousemove", h);
  }, []);
  return <div className="global-spotlight" />;
}

/* ──────────────────────────────────────────────────────────
   Navigation Config
   ────────────────────────────────────────────────────────── */
const NAV_USER = [
  { to: "/", label: "Dashboard", icon: LayoutDashboard },
  { to: "/upload", label: "Upload", icon: Upload },
  { to: "/groups", label: "Groups", icon: Users },
  { to: "/approvals", label: "Approvals", icon: CheckCircle2 },
  { to: "/verify", label: "Verify", icon: FileCheck2 },
];
const NAV_ADMIN = [
  { to: "/admin/users", label: "Users", icon: UserPlus2 },
  { to: "/admin/hods", label: "HOD Matrix", icon: Grid3X3 },
  { to: "/admin/analytics", label: "Analytics", icon: Activity },
  { to: "/admin/audit", label: "Audit Log", icon: BookOpen },
];

/* ──────────────────────────────────────────────────────────
   Confirm Dialog
   ────────────────────────────────────────────────────────── */
function ConfirmDialog({
  open,
  title,
  description,
  confirmLabel,
  onCancel,
  onConfirm,
}: {
  open: boolean;
  title: string;
  description: string;
  confirmLabel: string;
  onCancel: () => void;
  onConfirm: () => void | Promise<void>;
}) {
  return (
    <Dialog
      open={open}
      onOpenChange={(v) => {
        if (!v) onCancel();
      }}
    >
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle>{title}</DialogTitle>
          <DialogDescription>{description}</DialogDescription>
        </DialogHeader>
        <DialogFooter className="gap-2 sm:gap-0">
          <Button variant="outline" onClick={onCancel}>
            Cancel
          </Button>
          <Button variant="destructive" onClick={() => void onConfirm()}>
            {confirmLabel}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

/* ──────────────────────────────────────────────────────────
   Stat Card
   ────────────────────────────────────────────────────────── */
function Stat({
  title,
  value,
  icon,
  accent,
}: {
  title: string;
  value: string;
  icon: React.ReactNode;
  accent?: string;
}) {
  return (
    <div className="stat-card glass rounded-xl p-5 relative">
      <div className="flex items-center justify-between">
        <div className="space-y-1">
          <p className="text-xs font-medium text-muted-foreground uppercase tracking-wider">
            {title}
          </p>
          <p className="text-3xl font-bold tracking-tight">{value}</p>
        </div>
        <div
          className={`h-11 w-11 rounded-xl flex items-center justify-center ${
            accent || "bg-primary/8 text-primary"
          }`}
        >
          {icon}
        </div>
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
   Loading Screen
   ────────────────────────────────────────────────────────── */
function LoadingScreen() {
  return (
    <div className="min-h-screen bg-page-gradient flex items-center justify-center p-6">
      <motion.div
        initial={{ opacity: 0, scale: 0.96 }}
        animate={{ opacity: 1, scale: 1 }}
        className="max-w-md w-full text-center space-y-6"
      >
        <div className="space-y-2">
          <h1 className="font-display text-5xl text-primary">Premier DMS</h1>
          <p className="text-muted-foreground text-sm">
            Preparing your workspace...
          </p>
        </div>
        <div className="h-1.5 rounded-full bg-muted overflow-hidden mx-auto max-w-xs">
          <div className="h-full w-2/3 bg-primary/70 animate-shimmer rounded-full" />
        </div>
      </motion.div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   PROTECTED LAYOUT — Sidebar + Header + Content Area
   ══════════════════════════════════════════════════════════ */
function ProtectedLayout({ children }: { children: React.ReactNode }) {
  const { user, logout } = useAuth();
  const navigate = useNavigate();
  const location = useLocation();
  const [collapsed, setCollapsed] = useState(() => window.innerWidth < 1024);
  const [mobileOpen, setMobileOpen] = useState(false);

  const navItems = [...NAV_USER, ...(user?.isAdmin ? NAV_ADMIN : [])];
  const currentLabel =
    navItems.find((n) => n.to === location.pathname)?.label || "DMS";

  // Close mobile on navigate
  useEffect(() => {
    setMobileOpen(false);
  }, [location.pathname]);

  return (
    <div className="min-h-screen bg-page-gradient relative flex">
      <MouseGlow />
      <div className="dot-pattern fixed inset-0 pointer-events-none z-0" />

      {/* ── Desktop Sidebar ─────────────────────────────── */}
      <aside
        className={`hidden lg:flex flex-col fixed top-0 left-0 h-screen z-30 border-r border-border/60 bg-white/80 backdrop-blur-xl transition-all duration-300 ${
          collapsed ? "w-[var(--sidebar-collapsed-w)]" : "w-[var(--sidebar-w)]"
        }`}
      >
        {/* Brand */}
        <div className="h-[var(--header-h)] flex items-center gap-3 px-4 border-b border-border/40 shrink-0">
          <img
            src="/l.png"
            alt="Logo"
            className="h-8 w-8 object-contain shrink-0"
          />
          {!collapsed && (
            <motion.span
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="font-display text-2xl text-primary leading-none whitespace-nowrap"
            >
              Premier DMS
            </motion.span>
          )}
        </div>

        {/* Nav Links */}
        <nav className="flex-1 overflow-y-auto py-4 px-3 space-y-1">
          {!collapsed && (
            <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-widest px-3 mb-2">
              Navigation
            </p>
          )}
          {NAV_USER.map((item) => (
            <NavLink
              key={item.to}
              to={item.to}
              className={({ isActive }) =>
                `sidebar-link ${isActive ? "active" : ""} ${
                  collapsed ? "justify-center px-0" : ""
                }`
              }
              title={item.label}
            >
              <item.icon className="h-[18px] w-[18px] shrink-0" />
              {!collapsed && <span>{item.label}</span>}
            </NavLink>
          ))}

          {user?.isAdmin && (
            <>
              <div
                className={`my-3 ${
                  collapsed ? "mx-2" : "mx-3"
                } h-px bg-border/60`}
              />
              {!collapsed && (
                <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-widest px-3 mb-2">
                  Admin
                </p>
              )}
              {NAV_ADMIN.map((item) => (
                <NavLink
                  key={item.to}
                  to={item.to}
                  className={({ isActive }) =>
                    `sidebar-link ${isActive ? "active" : ""} ${
                      collapsed ? "justify-center px-0" : ""
                    }`
                  }
                  title={item.label}
                >
                  <item.icon className="h-[18px] w-[18px] shrink-0" />
                  {!collapsed && <span>{item.label}</span>}
                </NavLink>
              ))}
            </>
          )}
        </nav>

        {/* Collapse Toggle + User */}
        <div className="border-t border-border/40 p-3 space-y-2 shrink-0">
          <button
            onClick={() => setCollapsed((c) => !c)}
            className="w-full flex items-center justify-center gap-2 py-2 rounded-lg text-xs text-muted-foreground hover:bg-secondary/60 transition-colors"
          >
            {collapsed ? (
              <PanelLeftOpen className="h-4 w-4" />
            ) : (
              <>
                <PanelLeftClose className="h-4 w-4" />
                <span>Collapse</span>
              </>
            )}
          </button>
          {!collapsed && (
            <div className="flex items-center gap-2 px-1">
              <div className="h-8 w-8 rounded-full bg-primary/10 text-primary flex items-center justify-center text-xs font-bold shrink-0">
                {(user?.empName || user?.email || "U")[0].toUpperCase()}
              </div>
              <div className="min-w-0 flex-1">
                <p className="text-xs font-medium truncate">
                  {user?.empName || user?.email}
                </p>
                <p className="text-[10px] text-muted-foreground truncate">
                  {user?.department}
                </p>
              </div>
              <button
                onClick={logout}
                className="text-muted-foreground hover:text-destructive transition-colors p-1"
                title="Logout"
              >
                <LogOut className="h-3.5 w-3.5" />
              </button>
            </div>
          )}
        </div>
      </aside>

      {/* ── Mobile Menu ─────────────────────────────────── */}
      <AnimatePresence>
        {mobileOpen && (
          <>
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-black/40 backdrop-blur-sm z-40 lg:hidden"
              onClick={() => setMobileOpen(false)}
            />
            <motion.aside
              initial={{ x: "-100%" }}
              animate={{ x: 0 }}
              exit={{ x: "-100%" }}
              transition={{ type: "spring", damping: 28, stiffness: 300 }}
              className="fixed top-0 left-0 bottom-0 w-72 z-50 lg:hidden mobile-menu-backdrop text-white overflow-y-auto"
            >
              <div className="p-5 space-y-6">
                <div className="flex items-center justify-between">
                  <h2 className="font-display text-3xl">Menu</h2>
                  <button
                    onClick={() => setMobileOpen(false)}
                    className="p-1.5 rounded-lg bg-white/10 hover:bg-white/20 transition-colors"
                  >
                    <X className="h-5 w-5" />
                  </button>
                </div>
                <nav className="space-y-1">
                  {navItems.map((item, i) => (
                    <motion.div
                      key={item.to}
                      initial={{ opacity: 0, x: -12 }}
                      animate={{ opacity: 1, x: 0 }}
                      transition={{ delay: i * 0.04 }}
                    >
                      <NavLink
                        to={item.to}
                        onClick={() => setMobileOpen(false)}
                        className={({ isActive }) =>
                          `flex items-center gap-3 px-4 py-3 rounded-xl text-sm transition-all ${
                            isActive
                              ? "bg-white/20 font-semibold"
                              : "text-white/80 hover:bg-white/10"
                          }`
                        }
                      >
                        <item.icon className="h-4 w-4" />
                        {item.label}
                      </NavLink>
                    </motion.div>
                  ))}
                </nav>
                <div className="pt-4 border-t border-white/20 space-y-3 text-sm">
                  <p className="text-white/70">{user?.email}</p>
                  <Button
                    className="w-full bg-white/15 hover:bg-white/25 border-white/30 text-white"
                    variant="outline"
                    onClick={logout}
                  >
                    <LogOut className="h-4 w-4 mr-2" /> Sign Out
                  </Button>
                </div>
              </div>
            </motion.aside>
          </>
        )}
      </AnimatePresence>

      {/* ── Main Area ───────────────────────────────────── */}
      <div
        className={`flex-1 flex flex-col min-h-screen transition-all duration-300 ${
          collapsed
            ? "lg:ml-[var(--sidebar-collapsed-w)]"
            : "lg:ml-[var(--sidebar-w)]"
        }`}
      >
        {/* Top Bar */}
        <header className="sticky top-0 z-20 h-[var(--header-h)] flex items-center justify-between gap-4 px-4 md:px-6 border-b border-border/50 bg-white/70 backdrop-blur-xl">
          <div className="flex items-center gap-3">
            <button
              onClick={() => setMobileOpen(true)}
              className="lg:hidden p-2 rounded-lg hover:bg-secondary transition-colors"
            >
              <Menu className="h-5 w-5" />
            </button>
            <div className="flex items-center gap-2 lg:hidden">
              <img src="/l.png" alt="Logo" className="h-7 w-7" />
              <span className="font-display text-xl text-primary">DMS</span>
            </div>
            <h1 className="hidden lg:block text-lg font-semibold text-foreground">
              {currentLabel}
            </h1>
          </div>
          <div className="flex items-center gap-2">
            {user?.department && (
              <Badge variant="outline" className="hidden sm:flex text-xs">
                {user.department}
              </Badge>
            )}
            {user?.location && (
              <Badge variant="secondary" className="hidden md:flex text-xs">
                {user.location}
              </Badge>
            )}
            <div className="hidden lg:flex h-8 w-8 rounded-full bg-primary/10 text-primary items-center justify-center text-xs font-bold">
              {(user?.empName || user?.email || "U")[0].toUpperCase()}
            </div>
          </div>
        </header>

        {/* Page Content */}
        <main className="flex-1 p-4 md:p-6 lg:p-8 relative z-10">
          {children}
        </main>

        {/* Footer */}
        <footer className="border-t border-border/40 bg-white/50 backdrop-blur-sm px-4 md:px-6 lg:px-8 py-3">
          <div className="flex items-center justify-between text-xs text-muted-foreground gap-4">
            <div className="flex items-center gap-2">
              <img src="/l.png" alt="Logo" className="h-5 w-5 opacity-60" />
              <span className="hidden sm:inline">
                Premier Energies · Document Management System
              </span>
              <span className="sm:hidden">Premier DMS</span>
            </div>
            <span>© {new Date().getFullYear()}</span>
          </div>
        </footer>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   DASHBOARD
   ══════════════════════════════════════════════════════════ */
function DashboardPage() {
  const { user } = useAuth();
  const toast = useAppToast();
  const [docs, setDocs] = useState<Doc[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(false);
  const [q, setQ] = useState("");
  const [ctrl, setCtrl] = useState("all");
  const [dept, setDept] = useState("");
  const [loc, setLoc] = useState("");
  const [exp, setExp] = useState("all");
  const [view, setView] = useState<"grid" | "list">("grid");
  const [selected, setSelected] = useState<Doc | null>(null);
  const [history, setHistory] = useState<any[]>([]);
  const [perm, setPerm] = useState<any>(null);
  const [deleteDoc, setDeleteDoc] = useState<Doc | null>(null);
  const [deleting, setDeleting] = useState(false);
  const viewerRef = useRef<HTMLDivElement>(null);

  const isOwner = (d: Doc) =>
    d.CreatorEmail?.toLowerCase() === user?.email?.toLowerCase();

  const fetchDocs = useCallback(async () => {
    setLoading(true);
    try {
      const p = new URLSearchParams();
      if (q) p.set("search", q);
      if (dept) p.set("department", dept);
      if (loc) p.set("location", loc);
      if (ctrl !== "all") p.set("isControlled", String(ctrl === "controlled"));
      if (exp !== "all") p.set("expired", String(exp === "expired"));
      p.set("pageSize", "40");
      const data = await api<{ documents: Doc[]; total: number }>(
        `/api/documents?${p}`
      );
      setDocs(data.documents || []);
      setTotal(data.total || 0);
    } catch {
      /* handled by auth */
    } finally {
      setLoading(false);
    }
  }, [q, dept, loc, ctrl, exp]);

  useEffect(() => {
    if (user?.email) fetchDocs();
  }, [user?.email]);

  const openDoc = async (doc: Doc) => {
    setSelected(doc);
    try {
      const [logs, permission] = await Promise.all([
        api<{ logs: any[] }>(`/api/documents/${doc.Id}/audit`),
        api<any>(`/api/documents/${doc.Id}/permission`).catch(() => ({
          canPrint: false,
          canDownload: false,
        })),
      ]);
      setHistory(logs.logs || []);
      setPerm(permission);
    } catch {
      setHistory([]);
      setPerm(null);
    }
  };

  useEffect(() => {
    if (selected && viewerRef.current) {
      viewerRef.current.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }, [selected]);

  const confirmDelete = async () => {
    if (!deleteDoc || deleting) return;
    setDeleting(true);
    try {
      await api(`/api/documents/${deleteDoc.Id}`, {
        method: "DELETE",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ reason: "Deleted from dashboard" }),
      });
      setDocs((p) => p.filter((d) => d.Id !== deleteDoc.Id));
      setTotal((p) => Math.max(0, p - 1));
      if (selected?.Id === deleteDoc.Id) {
        setSelected(null);
        setHistory([]);
        setPerm(null);
      }
      toast({ title: "Document deleted", variant: "success" });
      setDeleteDoc(null);
    } catch (e) {
      toast({
        title: "Delete failed",
        description: String((e as Error)?.message || ""),
        variant: "error",
      });
    } finally {
      setDeleting(false);
    }
  };

  return (
    <div className="space-y-6 animate-fade-in">
      {/* Search */}
      <Card
        className="glass-elevated overflow-hidden"
        data-tour="dashboard-search"
      >
        <CardContent className="p-5">
          <div className="flex flex-col lg:flex-row gap-3">
            <div className="flex-1 relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
              <Input
                className="pl-10"
                placeholder="Search by title, content, metadata, creator..."
                value={q}
                onChange={(e) => setQ(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") fetchDocs();
                }}
              />
            </div>
            <div className="flex flex-wrap gap-2">
              <Input
                placeholder="Department"
                className="w-32"
                value={dept}
                onChange={(e) => setDept(e.target.value)}
              />
              <Input
                placeholder="Location"
                className="w-32"
                value={loc}
                onChange={(e) => setLoc(e.target.value)}
              />
              <Select value={ctrl} onValueChange={setCtrl}>
                <SelectTrigger className="w-36">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">All Types</SelectItem>
                  <SelectItem value="controlled">Controlled</SelectItem>
                  <SelectItem value="uncontrolled">Open</SelectItem>
                </SelectContent>
              </Select>
              <Select value={exp} onValueChange={setExp}>
                <SelectTrigger className="w-32">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">Any Validity</SelectItem>
                  <SelectItem value="expired">Expired</SelectItem>
                  <SelectItem value="active">Valid</SelectItem>
                </SelectContent>
              </Select>
              <Button onClick={fetchDocs} className="shrink-0">
                {loading ? (
                  <span className="flex items-center gap-2">
                    <span className="h-3 w-3 rounded-full border-2 border-white/30 border-t-white animate-spin" />
                    Searching
                  </span>
                ) : (
                  "Search"
                )}
              </Button>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Stats */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 stagger-children">
        <Stat
          title="Documents"
          value={String(total)}
          icon={<FileArchive className="h-5 w-5" />}
        />
        <Stat
          title="Controlled"
          value={String(docs.filter((d) => d.IsControlled).length)}
          icon={<ShieldCheck className="h-5 w-5" />}
          accent="bg-amber-500/10 text-amber-600"
        />
        <Stat
          title="Pending"
          value={String(
            docs.filter((d) => d.ApprovalStatus?.startsWith("pending")).length
          )}
          icon={<Clock3 className="h-5 w-5" />}
          accent="bg-orange-500/10 text-orange-600"
        />
        <Stat
          title="Approved"
          value={String(docs.filter((d) => d.Status === "approved").length)}
          icon={<CheckCircle2 className="h-5 w-5" />}
          accent="bg-emerald-500/10 text-emerald-600"
        />
      </div>

      {/* View Toggle + Results */}
      <div className="flex items-center justify-between">
        <p className="text-sm text-muted-foreground">
          {total} document{total !== 1 ? "s" : ""}
        </p>
        <div className="flex items-center gap-1 border rounded-lg p-0.5">
          <button
            onClick={() => setView("grid")}
            className={`p-1.5 rounded-md transition-colors ${
              view === "grid"
                ? "bg-primary text-white"
                : "text-muted-foreground hover:text-foreground"
            }`}
          >
            <Grid3X3 className="h-4 w-4" />
          </button>
          <button
            onClick={() => setView("list")}
            className={`p-1.5 rounded-md transition-colors ${
              view === "list"
                ? "bg-primary text-white"
                : "text-muted-foreground hover:text-foreground"
            }`}
          >
            <List className="h-4 w-4" />
          </button>
        </div>
      </div>

      {/* Document Cards */}
      <div
        className={`stagger-children ${
          view === "grid"
            ? "grid md:grid-cols-2 xl:grid-cols-3 gap-4"
            : "space-y-3"
        }`}
      >
        {docs.map((doc, i) => (
          <div
            key={doc.Id}
            className={`doc-card glass rounded-xl cursor-pointer ${
              view === "list" ? "flex items-center gap-4 p-4" : "p-5 space-y-3"
            }`}
            data-tour={i === 0 ? "dashboard-doc-card" : undefined}
            onClick={() => openDoc(doc)}
          >
            {view === "list" ? (
              /* List view */
              <>
                <div
                  className={`h-10 w-10 rounded-lg flex items-center justify-center shrink-0 ${
                    doc.IsControlled
                      ? "bg-amber-500/10 text-amber-600"
                      : "bg-primary/8 text-primary"
                  }`}
                >
                  <FileText className="h-5 w-5" />
                </div>
                <div className="flex-1 min-w-0">
                  <h3 className="font-semibold text-sm truncate">
                    {doc.Title}
                  </h3>
                  <p className="text-xs text-muted-foreground truncate">
                    {doc.FileName} · {doc.Department} ·{" "}
                    {new Date(doc.CreatedAt).toLocaleDateString()}
                  </p>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <Badge variant="outline" className="text-[10px]">
                    v{doc.CurrentVersionLabel || doc.CurrentVersion}
                  </Badge>
                  <Badge
                    variant={doc.IsControlled ? "warning" : "secondary"}
                    className="text-[10px]"
                  >
                    {doc.IsControlled ? "Controlled" : "Open"}
                  </Badge>
                </div>
              </>
            ) : (
              /* Grid view */
              <>
                <div className="flex items-start justify-between gap-2">
                  <div
                    className={`h-9 w-9 rounded-lg flex items-center justify-center shrink-0 ${
                      doc.IsControlled
                        ? "bg-amber-500/10 text-amber-600"
                        : "bg-primary/8 text-primary"
                    }`}
                  >
                    <FileText className="h-4 w-4" />
                  </div>
                  <div className="flex gap-1.5 flex-wrap">
                    <Badge
                      variant={doc.IsControlled ? "warning" : "secondary"}
                      className="text-[10px]"
                    >
                      {doc.IsControlled ? "Controlled" : "Open"}
                    </Badge>
                    {doc.HodSkipped && (
                      <Badge variant="warning" className="text-[10px]">
                        HOD Skipped
                      </Badge>
                    )}
                  </div>
                </div>
                <div>
                  <h3 className="font-semibold text-sm leading-tight line-clamp-1">
                    {doc.Title}
                  </h3>
                  <p className="text-xs text-muted-foreground line-clamp-2 mt-1">
                    {doc.Description || doc.FileName}
                  </p>
                </div>
                <div className="flex items-center justify-between pt-1">
                  <div className="flex gap-1.5">
                    <Badge variant="outline" className="text-[10px]">
                      v{doc.CurrentVersionLabel || doc.CurrentVersion}
                    </Badge>
                    <Badge variant="info" className="text-[10px]">
                      {doc.ApprovalStatus || "none"}
                    </Badge>
                  </div>
                  <span className="text-[10px] text-muted-foreground">
                    {new Date(doc.CreatedAt).toLocaleDateString()}
                  </span>
                </div>
                <div className="flex gap-2 pt-1">
                  <Button
                    size="sm"
                    variant="secondary"
                    className="text-xs h-8"
                    onClick={(e) => {
                      e.stopPropagation();
                      openDoc(doc);
                    }}
                  >
                    <Eye className="h-3 w-3 mr-1" /> View
                  </Button>
                  {isOwner(doc) && (
                    <Button
                      size="sm"
                      variant="ghost"
                      className="text-xs h-8 text-destructive hover:text-destructive"
                      onClick={(e) => {
                        e.stopPropagation();
                        setDeleteDoc(doc);
                      }}
                    >
                      <Trash2 className="h-3 w-3" />
                    </Button>
                  )}
                </div>
              </>
            )}
          </div>
        ))}
      </div>

      {docs.length === 0 && !loading && (
        <div className="text-center py-16 space-y-3">
          <FolderOpen className="h-12 w-12 text-muted-foreground/40 mx-auto" />
          <p className="text-muted-foreground">
            No documents found. Try adjusting your filters.
          </p>
        </div>
      )}

      {/* Selected Document Viewer */}
      {selected && (
        <motion.div
          ref={viewerRef}
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          className="space-y-4"
        >
          <Card className="glass-elevated overflow-hidden">
            <CardHeader className="pb-3">
              <div className="flex items-start justify-between gap-4">
                <div>
                  <CardTitle className="text-xl">{selected.Title}</CardTitle>
                  <CardDescription className="mt-1">
                    {selected.FileName} · v
                    {selected.CurrentVersionLabel || selected.CurrentVersion}
                  </CardDescription>
                </div>
                <div className="flex gap-2">
                  {perm?.canDownload && (
                    <Button size="sm" variant="outline" asChild>
                      <a
                        href={`/api/documents/${selected.Id}/download`}
                        target="_blank"
                        rel="noreferrer"
                      >
                        <Download className="h-3.5 w-3.5 mr-1.5" />
                        Download
                      </a>
                    </Button>
                  )}
                  <Button
                    size="sm"
                    variant="ghost"
                    onClick={() => {
                      setSelected(null);
                      setHistory([]);
                      setPerm(null);
                    }}
                  >
                    <X className="h-4 w-4" />
                  </Button>
                </div>
              </div>
              {perm && (
                <div className="flex gap-1.5 mt-2 text-[10px]">
                  <Badge variant="outline">
                    <Eye className="h-2.5 w-2.5 mr-1" />
                    View
                  </Badge>
                  {perm.canPrint && (
                    <Badge variant="outline">
                      <Printer className="h-2.5 w-2.5 mr-1" />
                      Print
                    </Badge>
                  )}
                  {perm.canDownload && (
                    <Badge variant="outline">
                      <Download className="h-2.5 w-2.5 mr-1" />
                      Download
                    </Badge>
                  )}
                  {!perm.canPrint && !perm.canDownload && (
                    <Badge variant="secondary">
                      <Lock className="h-2.5 w-2.5 mr-1" />
                      View Only
                    </Badge>
                  )}
                </div>
              )}
            </CardHeader>
            <CardContent className="space-y-4">
              <div
                className="rounded-xl border overflow-hidden bg-slate-50"
                onContextMenu={(e) => e.preventDefault()}
                onKeyDown={(e) => {
                  if (
                    (e.ctrlKey || e.metaKey) &&
                    ["c", "p", "s"].includes(e.key.toLowerCase())
                  )
                    e.preventDefault();
                }}
              >
                <iframe
                  title="viewer"
                  className="w-full h-[480px]"
                  src={`/api/documents/${selected.Id}/view?embed=1#toolbar=0&navpanes=0&scrollbar=0`}
                />
              </div>

              {/* Public Links */}
              <PublicLinksPanel docId={selected.Id} />

              {/* History */}
              <div>
                <h4 className="text-sm font-semibold mb-2 flex items-center gap-2">
                  <BookOpen className="h-4 w-4 text-muted-foreground" />{" "}
                  Activity Timeline
                </h4>
                <div className="max-h-60 overflow-auto rounded-lg border divide-y">
                  {history.length === 0 && (
                    <p className="p-4 text-sm text-muted-foreground">
                      No activity recorded.
                    </p>
                  )}
                  {history.map((h, i) => (
                    <div
                      key={i}
                      className="px-4 py-3 text-sm hover:bg-muted/30 transition-colors"
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-medium text-xs">{h.Action}</span>
                        <span className="text-[10px] text-muted-foreground">
                          {new Date(h.CreatedAt).toLocaleString()}
                        </span>
                      </div>
                      <p className="text-xs text-muted-foreground mt-0.5">
                        {h.UserEmail || "System"}
                        {h.Reason ? ` — ${h.Reason}` : ""}
                      </p>
                    </div>
                  ))}
                </div>
              </div>
            </CardContent>
          </Card>
        </motion.div>
      )}

      <ConfirmDialog
        open={!!deleteDoc}
        title="Delete Document?"
        description={
          deleteDoc
            ? `"${deleteDoc.Title}" will be soft-deleted and removed from search.`
            : ""
        }
        confirmLabel={deleting ? "Deleting..." : "Delete"}
        onCancel={() => !deleting && setDeleteDoc(null)}
        onConfirm={confirmDelete}
      />
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
   Public Links Panel
   ────────────────────────────────────────────────────────── */
function PublicLinksPanel({ docId }: { docId: number }) {
  const toast = useAppToast();
  const [links, setLinks] = useState<any[]>([]);
  const [exp, setExp] = useState("");

  const load = () =>
    api<{ links: any[] }>(`/api/documents/${docId}/public-links`)
      .then((r) => setLinks(r.links || []))
      .catch(() => setLinks([]));
  useEffect(() => {
    load();
  }, [docId]);

  const create = async () => {
    await api("/api/public-links", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ docId, expiresAt: exp || null }),
    });
    toast({ title: "Public link created", variant: "success" });
    load();
  };

  return (
    <div className="space-y-2">
      <h4 className="text-sm font-semibold flex items-center gap-2">
        <Share2 className="h-4 w-4 text-muted-foreground" /> Public Links
      </h4>
      <div className="flex gap-2 flex-wrap">
        <Input
          type="datetime-local"
          value={exp}
          onChange={(e) => setExp(e.target.value)}
          className="max-w-[200px] text-xs h-9"
        />
        <Button size="sm" onClick={create} className="h-9">
          <Plus className="h-3 w-3 mr-1" />
          Generate
        </Button>
      </div>
      {links.map((l) => {
        const url = `${BASE}/public/${l.LinkToken}`;
        return (
          <div
            key={l.Id}
            className="flex items-center gap-2 p-2 rounded-lg border bg-muted/30 text-xs"
          >
            <a
              href={url}
              className="text-primary hover:underline truncate flex-1"
              target="_blank"
              rel="noreferrer"
            >
              {url}
            </a>
            <Button
              size="sm"
              variant="ghost"
              className="h-7 w-7 p-0"
              onClick={() => {
                navigator.clipboard.writeText(url);
                toast({ title: "Copied!", variant: "info" });
              }}
            >
              <Copy className="h-3 w-3" />
            </Button>
          </div>
        );
      })}
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   UPLOAD PAGE
   ══════════════════════════════════════════════════════════ */
function UploadPage() {
  const { user } = useAuth();
  const navigate = useNavigate();
  const toast = useAppToast();
  const [title, setTitle] = useState("");
  const [description, setDescription] = useState("");
  const [metadata, setMetadata] = useState("{}");
  const [isControlled, setIsControlled] = useState(false);
  const [validFrom, setValidFrom] = useState("");
  const [validTo, setValidTo] = useState("");
  const [shareScope, setShareScope] = useState("private");
  const [shareGroupId, setShareGroupId] = useState("");
  const [defaultAccessType, setDefaultAccessType] = useState("view_only");
  const [groups, setGroups] = useState<any[]>([]);
  const [reason, setReason] = useState("");
  const [existingDocs, setExistingDocs] = useState<Doc[]>([]);
  const [newVersionBaseId, setNewVersionBaseId] = useState("__new__");
  const [file, setFile] = useState<File | null>(null);
  const [saving, setSaving] = useState(false);
  const [progress, setProgress] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    api<any[]>("/api/share-groups")
      .then(setGroups)
      .catch(() => setGroups([]));
    api<{ documents: Doc[] }>("/api/documents?isControlled=true&pageSize=200")
      .then((r) => setExistingDocs(r.documents || []))
      .catch(() => setExistingDocs([]));
  }, []);

  const upload = async () => {
    if (!file) {
      toast({ title: "Select a file", variant: "info" });
      return;
    }
    if (isControlled && !reason.trim()) {
      toast({ title: "Revision reason required", variant: "info" });
      return;
    }
    if (!validFrom || !validTo) {
      toast({ title: "Validity dates required", variant: "info" });
      return;
    }

    setSaving(true);
    setProgress(0);
    try {
      const form = new FormData();
      form.append("file", file);
      form.append("title", title || file.name);
      form.append("description", description);
      form.append("metadata", metadata);
      form.append("isControlled", String(isControlled));
      form.append("shareScope", shareScope);
      if (shareScope === "group" && shareGroupId)
        form.append("shareGroupId", shareGroupId);
      form.append("defaultAccessType", defaultAccessType);
      if (reason) form.append("reason", reason);
      form.append("validFrom", validFrom);
      form.append("validTo", validTo);

      const url =
        newVersionBaseId !== "__new__"
          ? `/api/documents/${newVersionBaseId}/new-version`
          : "/api/documents/upload";

      await new Promise<void>((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        xhr.open("POST", url, true);
        xhr.withCredentials = true;
        xhr.upload.onprogress = (e) => {
          if (e.lengthComputable)
            setProgress(Math.round((e.loaded / e.total) * 95));
        };
        xhr.onload = () => {
          setProgress(100);
          xhr.status >= 200 && xhr.status < 300
            ? resolve()
            : reject(new Error(xhr.responseText));
        };
        xhr.onerror = () => reject(new Error("Network error"));
        xhr.send(form);
      });

      setTitle("");
      setDescription("");
      setIsControlled(false);
      setReason("");
      setFile(null);
      setValidFrom("");
      setValidTo("");
      setNewVersionBaseId("__new__");
      toast({ title: "Upload successful", variant: "success" });
    } catch (e) {
      toast({
        title: "Upload failed",
        description: String((e as Error)?.message || ""),
        variant: "error",
      });
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto space-y-6 animate-fade-in">
      <div>
        <h2 className="font-display text-3xl text-foreground">
          Upload Document
        </h2>
        <p className="text-sm text-muted-foreground mt-1">
          Upload with optional controlled workflow, versioning, and targeted
          sharing.
        </p>
      </div>

      <Card className="glass">
        <CardContent className="p-6 space-y-5">
          {/* File Drop Zone */}
          <div
            className="border-2 border-dashed border-border/80 rounded-xl p-8 text-center hover:border-primary/40 transition-colors cursor-pointer"
            onClick={() => fileInputRef.current?.click()}
          >
            <Upload className="h-8 w-8 text-muted-foreground/50 mx-auto mb-3" />
            {file ? (
              <p className="text-sm font-medium">
                {file.name}{" "}
                <span className="text-muted-foreground">
                  ({(file.size / 1024).toFixed(0)} KB)
                </span>
              </p>
            ) : (
              <p className="text-sm text-muted-foreground">
                Click to select a file{" "}
                <span className="text-xs">
                  (.doc, .docx, .xls, .xlsx, .pdf)
                </span>
              </p>
            )}
            <input
              ref={fileInputRef}
              type="file"
              className="hidden"
              accept=".doc,.docx,.xls,.xlsx,.pdf"
              onChange={(e) => setFile(e.target.files?.[0] || null)}
            />
          </div>

          {/* Controlled Toggle */}
          <div
            className="flex items-center gap-3 p-3 rounded-lg bg-amber-50/60 border border-amber-200/60"
            data-tour="upload-controlled"
          >
            <Checkbox
              checked={isControlled}
              onCheckedChange={(v) => setIsControlled(Boolean(v))}
              id="ctrl"
            />
            <Label htmlFor="ctrl" className="cursor-pointer">
              <span className="font-medium">Controlled Document</span>
              <span className="text-xs text-muted-foreground ml-2">
                Requires RM → HOD → DC approval
              </span>
            </Label>
          </div>

          {/* Form Fields */}
          <div className="grid md:grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <Label>Title</Label>
              <Input
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                placeholder="Document title"
              />
            </div>
            {isControlled && (
              <div className="space-y-1.5">
                <Label>Revision Reason *</Label>
                <Input
                  value={reason}
                  onChange={(e) => setReason(e.target.value)}
                  placeholder="Why is this revision needed?"
                />
              </div>
            )}
          </div>

          <div className="grid md:grid-cols-3 gap-4">
            <div className="space-y-1.5">
              <Label>Version Against</Label>
              <Select
                value={newVersionBaseId}
                onValueChange={setNewVersionBaseId}
              >
                <SelectTrigger>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="__new__">New Document</SelectItem>
                  {existingDocs.map((d) => (
                    <SelectItem key={d.Id} value={String(d.Id)}>
                      {d.Title} (v{d.CurrentVersionLabel || d.CurrentVersion})
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="space-y-1.5">
              <Label>Valid From *</Label>
              <Input
                type="date"
                value={validFrom}
                onChange={(e) => setValidFrom(e.target.value)}
              />
            </div>
            <div className="space-y-1.5">
              <Label>Valid To *</Label>
              <Input
                type="date"
                value={validTo}
                onChange={(e) => setValidTo(e.target.value)}
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <Label>Description</Label>
            <textarea
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              rows={3}
              className="w-full rounded-lg border border-input bg-background px-3 py-2 text-sm focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring"
            />
          </div>

          <div className="grid md:grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <Label>Share Scope</Label>
              <Select value={shareScope} onValueChange={setShareScope}>
                <SelectTrigger>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="private">Private</SelectItem>
                  <SelectItem value="group">Custom Group</SelectItem>
                  <SelectItem value="department">Department</SelectItem>
                  <SelectItem value="company">Company</SelectItem>
                </SelectContent>
              </Select>
            </div>
            <div className="space-y-1.5">
              <Label>Default Permission</Label>
              <Select
                value={defaultAccessType}
                onValueChange={setDefaultAccessType}
              >
                <SelectTrigger>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="view_only">View Only</SelectItem>
                  <SelectItem value="view_print">View + Print</SelectItem>
                  <SelectItem value="view_print_download">
                    View + Print + Download
                  </SelectItem>
                </SelectContent>
              </Select>
            </div>
          </div>

          {shareScope === "group" && (
            <div className="space-y-1.5">
              <Label>Select Group</Label>
              <Select value={shareGroupId} onValueChange={setShareGroupId}>
                <SelectTrigger>
                  <SelectValue placeholder="Choose group" />
                </SelectTrigger>
                <SelectContent>
                  {groups.map((g) => (
                    <SelectItem key={g.Id} value={String(g.Id)}>
                      {g.Name}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
          )}

          <Button
            onClick={upload}
            disabled={saving}
            className="w-full h-12 text-base"
            data-tour="upload-submit"
          >
            {saving ? (
              <span className="flex items-center gap-2">
                <span className="h-4 w-4 rounded-full border-2 border-white/30 border-t-white animate-spin" />
                {progress}%
              </span>
            ) : (
              "Upload Document"
            )}
          </Button>
        </CardContent>
      </Card>

      {/* Workflow Info */}
      <Card className="glass">
        <CardContent className="p-5">
          <div className="flex items-start gap-3">
            <div className="h-9 w-9 rounded-lg bg-amber-500/10 text-amber-600 flex items-center justify-center shrink-0 mt-0.5">
              <Shield className="h-4 w-4" />
            </div>
            <div className="text-sm space-y-1">
              <p className="font-semibold">Controlled Workflow</p>
              <p className="text-muted-foreground">
                Reporting Manager → HOD → Document Controller. HOD auto-skips if
                no mapping exists. Every action is immutably logged.
              </p>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   GROUPS PAGE
   ══════════════════════════════════════════════════════════ */
function GroupsPage() {
  const toast = useAppToast();
  const [groups, setGroups] = useState<any[]>([]);
  const [name, setName] = useState("");
  const [search, setSearch] = useState("");
  const [emps, setEmps] = useState<Employee[]>([]);
  const [members, setMembers] = useState<string[]>([]);
  const [editId, setEditId] = useState<number | null>(null);

  const load = () =>
    api<any[]>("/api/share-groups")
      .then(setGroups)
      .catch(() => setGroups([]));
  useEffect(() => {
    load();
  }, []);

  useEffect(() => {
    if (!search.trim()) {
      setEmps([]);
      return;
    }
    const t = setTimeout(() => {
      api<Employee[]>(`/api/employees/search?q=${encodeURIComponent(search)}`)
        .then(setEmps)
        .catch(() => setEmps([]));
    }, 200);
    return () => clearTimeout(t);
  }, [search]);

  const save = async () => {
    if (!name.trim() || !members.length) {
      toast({ title: "Name and members required", variant: "info" });
      return;
    }
    if (editId) {
      await api(`/api/share-groups/${editId}`, {
        method: "PUT",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ name, members }),
      });
      toast({ title: "Group updated", variant: "success" });
    } else {
      await api("/api/share-groups", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ name, members }),
      });
      toast({ title: "Group created", variant: "success" });
    }
    setName("");
    setMembers([]);
    setEditId(null);
    load();
  };

  return (
    <div className="max-w-5xl mx-auto space-y-6 animate-fade-in">
      <div>
        <h2 className="font-display text-3xl">Sharing Groups</h2>
        <p className="text-sm text-muted-foreground mt-1">
          Create reusable groups for document sharing.
        </p>
      </div>

      <div className="grid lg:grid-cols-5 gap-6">
        <Card className="lg:col-span-3 glass">
          <CardContent className="p-6 space-y-4">
            <Input
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="Group name"
            />
            <Input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search employees..."
            />
            <div className="max-h-32 overflow-auto border rounded-lg divide-y text-sm">
              {emps.map((e) => (
                <button
                  key={e.EmpEmail}
                  type="button"
                  className="w-full text-left p-2.5 hover:bg-muted/50 transition-colors"
                  onClick={() =>
                    setMembers((p) => [
                      ...new Set([...p, e.EmpEmail.toLowerCase()]),
                    ])
                  }
                >
                  <span className="font-medium">{e.EmpName}</span>{" "}
                  <span className="text-muted-foreground">({e.EmpEmail})</span>
                </button>
              ))}
            </div>
            <div className="flex flex-wrap gap-1.5 min-h-[32px]">
              {members.map((m) => (
                <button
                  key={m}
                  className="text-xs px-2.5 py-1 rounded-full border bg-muted/50 hover:bg-destructive/10 hover:border-destructive/30 transition-colors"
                  onClick={() => setMembers((p) => p.filter((x) => x !== m))}
                >
                  {m} ×
                </button>
              ))}
            </div>
            <div className="flex gap-2">
              <Button onClick={save}>
                {editId ? "Update" : "Create"} Group
              </Button>
              {editId && (
                <Button
                  variant="outline"
                  onClick={() => {
                    setEditId(null);
                    setName("");
                    setMembers([]);
                  }}
                >
                  Cancel
                </Button>
              )}
            </div>
          </CardContent>
        </Card>

        <Card className="lg:col-span-2 glass">
          <CardHeader className="pb-3">
            <CardTitle className="text-base">Your Groups</CardTitle>
          </CardHeader>
          <CardContent className="max-h-[500px] overflow-auto space-y-2">
            {groups.map((g) => (
              <div
                key={g.Id}
                className="p-3 border rounded-lg text-sm space-y-2 hover:bg-muted/20 transition-colors"
              >
                <div className="font-semibold">{g.Name}</div>
                <div className="text-xs text-muted-foreground">
                  {(g.members || []).length} members
                </div>
                <div className="flex gap-2">
                  <Button
                    size="sm"
                    variant="outline"
                    className="h-7 text-xs"
                    onClick={() => {
                      setEditId(g.Id);
                      setName(g.Name);
                      setMembers(g.members || []);
                      window.scrollTo({ top: 0, behavior: "smooth" });
                    }}
                  >
                    Edit
                  </Button>
                  <Button
                    size="sm"
                    variant="ghost"
                    className="h-7 text-xs text-destructive"
                    onClick={async () => {
                      await api(`/api/share-groups/${g.Id}`, {
                        method: "DELETE",
                      });
                      toast({ title: "Deleted", variant: "success" });
                      load();
                    }}
                  >
                    Delete
                  </Button>
                </div>
              </div>
            ))}
            {groups.length === 0 && (
              <p className="text-sm text-muted-foreground py-4 text-center">
                No groups yet.
              </p>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   APPROVALS PAGE
   ══════════════════════════════════════════════════════════ */
function ApprovalsPage() {
  const [pending, setPending] = useState<Approval[]>([]);
  const [comments, setComments] = useState<Record<number, string>>({});
  const [viewId, setViewId] = useState<number | null>(null);
  const toast = useAppToast();

  const load = () =>
    api<Approval[]>("/api/approvals/pending")
      .then(setPending)
      .catch(() => setPending([]));
  useEffect(() => {
    load();
  }, []);

  const decide = async (id: number, action: "approve" | "reject") => {
    await api(`/api/approvals/${id}/${action}`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ comments: comments[id] || "" }),
    });
    toast({
      title: action === "approve" ? "Approved" : "Rejected",
      variant: action === "approve" ? "success" : "info",
    });
    load();
  };

  return (
    <div
      className="max-w-4xl mx-auto space-y-6 animate-fade-in"
      data-tour="approvals-list"
    >
      <div>
        <h2 className="font-display text-3xl">Pending Approvals</h2>
        <p className="text-sm text-muted-foreground mt-1">
          {pending.length} item{pending.length !== 1 ? "s" : ""} awaiting your
          action.
        </p>
      </div>

      {pending.length === 0 && (
        <div className="text-center py-16">
          <CheckCircle2 className="h-12 w-12 text-emerald-400/60 mx-auto mb-3" />
          <p className="text-muted-foreground">
            All caught up — no pending approvals.
          </p>
        </div>
      )}

      <div className="space-y-3 stagger-children">
        {pending.map((p) => (
          <Card key={p.Id} className="glass overflow-hidden">
            <CardContent className="p-5 space-y-3">
              <div className="flex items-start justify-between gap-3 flex-wrap">
                <div>
                  <h3 className="font-semibold">
                    {p.Title}{" "}
                    <span className="text-muted-foreground font-normal">
                      v{p.CurrentVersion}
                    </span>
                  </h3>
                  <p className="text-xs text-muted-foreground mt-0.5">
                    Stage: {p.Stage.replace(/_/g, " ")} · From: {p.CreatorEmail}
                  </p>
                </div>
                <div className="flex gap-1.5">
                  {p.HodSkipped && (
                    <Badge variant="warning" className="text-[10px]">
                      <AlertTriangle className="h-2.5 w-2.5 mr-1" />
                      HOD Skipped
                    </Badge>
                  )}
                  <Badge variant="info" className="text-[10px]">
                    {p.Department} · {p.Location}
                  </Badge>
                </div>
              </div>
              <Input
                placeholder="Add comments (optional)"
                value={comments[p.Id] || ""}
                onChange={(e) =>
                  setComments((c) => ({ ...c, [p.Id]: e.target.value }))
                }
              />
              <div className="flex gap-2 flex-wrap">
                <Button
                  size="sm"
                  onClick={() => decide(p.Id, "approve")}
                  className="bg-emerald-600 hover:bg-emerald-700"
                >
                  <CheckCircle2 className="h-3.5 w-3.5 mr-1.5" />
                  Approve
                </Button>
                <Button
                  size="sm"
                  variant="destructive"
                  onClick={() => decide(p.Id, "reject")}
                >
                  Reject
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => setViewId(p.DocId)}
                >
                  <Eye className="h-3.5 w-3.5 mr-1.5" />
                  Preview
                </Button>
              </div>
            </CardContent>
          </Card>
        ))}
      </div>

      <Dialog
        open={!!viewId}
        onOpenChange={(v) => {
          if (!v) setViewId(null);
        }}
      >
        <DialogContent className="max-w-5xl">
          <DialogHeader>
            <DialogTitle>Document Preview</DialogTitle>
          </DialogHeader>
          {viewId && (
            <iframe
              title="preview"
              className="w-full h-[70vh] rounded-xl border"
              src={`/api/documents/${viewId}/view?embed=1`}
            />
          )}
          <DialogFooter>
            <Button variant="outline" onClick={() => setViewId(null)}>
              Close
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   VERIFY PAGE
   ══════════════════════════════════════════════════════════ */
function VerifyPage() {
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<any>(null);
  const [loading, setLoading] = useState(false);

  const verify = async () => {
    if (!file) return;
    setLoading(true);
    const form = new FormData();
    form.append("file", file);
    const data = await api<any>("/api/documents/verify-upload", {
      method: "POST",
      body: form,
    });
    setResult(data);
    setLoading(false);
  };

  return (
    <div
      className="max-w-2xl mx-auto space-y-6 animate-fade-in"
      data-tour="verify-panel"
    >
      <div>
        <h2 className="font-display text-3xl">Verify Controlled Copy</h2>
        <p className="text-sm text-muted-foreground mt-1">
          Upload a file to check if it matches an official controlled document.
        </p>
      </div>
      <Card className="glass">
        <CardContent className="p-6 space-y-4">
          <Input
            type="file"
            onChange={(e) => setFile(e.target.files?.[0] || null)}
          />
          <Button onClick={verify} disabled={!file || loading}>
            {loading ? "Verifying..." : "Verify"}
          </Button>
          {result && (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className={`p-4 rounded-xl border ${
                result.matched
                  ? "bg-emerald-50 border-emerald-200"
                  : "bg-red-50 border-red-200"
              }`}
            >
              <div className="flex items-center gap-2 mb-2">
                {result.matched ? (
                  <CheckCircle2 className="h-5 w-5 text-emerald-600" />
                ) : (
                  <AlertTriangle className="h-5 w-5 text-red-600" />
                )}
                <span className="font-semibold text-sm">
                  {result.matched ? "Match Found" : "No Match"}
                </span>
              </div>
              <p className="text-xs font-mono break-all text-muted-foreground">
                Hash: {result.uploadHash}
              </p>
              {result.documents?.[0] && (
                <p className="text-sm mt-2">
                  Document: <strong>{result.documents[0].Title}</strong> v
                  {result.documents[0].CurrentVersion}
                </p>
              )}
            </motion.div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   ADMIN: Users
   ══════════════════════════════════════════════════════════ */
function AdminUsersPage() {
  const toast = useAppToast();
  const [users, setUsers] = useState<Employee[]>([]);
  const [admins, setAdmins] = useState<any[]>([]);
  const [search, setSearch] = useState("");
  const [newAdmin, setNewAdmin] = useState("");
  const [newUser, setNewUser] = useState({
    EmpID: "",
    EmpName: "",
    EmpEmail: "",
    Department: "",
    Location: "",
    ReportingManagerID: "",
  });

  const load = async () => {
    const [u, a] = await Promise.all([
      api<{ users: Employee[] }>(
        `/api/admin/users?pageSize=100&search=${encodeURIComponent(search)}`
      ),
      api<any[]>("/api/admin/admins"),
    ]);
    setUsers(u.users || []);
    setAdmins(a || []);
  };
  useEffect(() => {
    load().catch(() => undefined);
  }, [search]);

  return (
    <div className="space-y-6 animate-fade-in">
      <div>
        <h2 className="font-display text-3xl">User Management</h2>
        <p className="text-sm text-muted-foreground mt-1">
          Create, search, and manage DMS users and admins.
        </p>
      </div>

      {/* Create User */}
      <Card className="glass" data-tour="users-create">
        <CardHeader className="pb-3">
          <CardTitle className="text-base flex items-center gap-2">
            <UserPlus2 className="h-4 w-4" />
            Create User
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="grid md:grid-cols-3 gap-3">
            {Object.entries(newUser).map(([k, v]) => (
              <Input
                key={k}
                placeholder={k}
                value={v}
                onChange={(e) =>
                  setNewUser((p) => ({ ...p, [k]: e.target.value }))
                }
              />
            ))}
          </div>
          <Button
            className="mt-3"
            onClick={async () => {
              await api("/api/admin/users", {
                method: "POST",
                headers: { "content-type": "application/json" },
                body: JSON.stringify(newUser),
              });
              setNewUser({
                EmpID: "",
                EmpName: "",
                EmpEmail: "",
                Department: "",
                Location: "",
                ReportingManagerID: "",
              });
              toast({ title: "User created", variant: "success" });
              load();
            }}
          >
            Create
          </Button>
        </CardContent>
      </Card>

      <div className="grid xl:grid-cols-3 gap-6">
        {/* User List */}
        <Card className="xl:col-span-2 glass">
          <CardHeader className="pb-3">
            <CardTitle className="text-base">Employee Directory</CardTitle>
          </CardHeader>
          <CardContent className="space-y-3">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
              <Input
                className="pl-10"
                placeholder="Search users..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            </div>
            <div className="max-h-[500px] overflow-auto rounded-lg border divide-y">
              {users.map((u) => (
                <div
                  key={u.EmpEmail}
                  className="p-3 hover:bg-muted/20 transition-colors"
                >
                  <div className="flex items-center justify-between gap-2">
                    <div>
                      <p className="text-sm font-medium">
                        {u.EmpName}{" "}
                        <span className="text-muted-foreground font-normal">
                          ({u.EmpID})
                        </span>
                      </p>
                      <p className="text-xs text-muted-foreground">
                        {u.EmpEmail} · {u.Department || u.Dept} ·{" "}
                        {u.Location || u.EmpLocation}
                      </p>
                    </div>
                    <Button
                      size="sm"
                      variant="ghost"
                      className="text-xs text-destructive h-7"
                      onClick={async () => {
                        await api(
                          `/api/admin/users/${encodeURIComponent(u.EmpEmail)}`,
                          { method: "DELETE" }
                        );
                        toast({
                          title: "User deactivated",
                          variant: "success",
                        });
                        load();
                      }}
                    >
                      Deactivate
                    </Button>
                  </div>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>

        {/* Admins */}
        <Card className="glass">
          <CardHeader className="pb-3">
            <CardTitle className="text-base">Admin Users</CardTitle>
          </CardHeader>
          <CardContent className="space-y-3">
            <div className="flex gap-2">
              <Input
                placeholder="admin@email.com"
                value={newAdmin}
                onChange={(e) => setNewAdmin(e.target.value)}
                className="text-sm"
              />
              <Button
                size="sm"
                onClick={async () => {
                  await api("/api/admin/admins", {
                    method: "POST",
                    headers: { "content-type": "application/json" },
                    body: JSON.stringify({ email: newAdmin }),
                  });
                  setNewAdmin("");
                  toast({ title: "Admin added", variant: "success" });
                  load();
                }}
              >
                Add
              </Button>
            </div>
            <div className="max-h-[400px] overflow-auto rounded-lg border divide-y">
              {admins.map((a) => (
                <div
                  key={a.Email}
                  className="p-2.5 flex items-center justify-between text-sm"
                >
                  <span className="truncate">{a.Email}</span>
                  <Badge
                    variant={a.Active ? "success" : "secondary"}
                    className="text-[10px] shrink-0"
                  >
                    {a.Active ? "Active" : "Off"}
                  </Badge>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   ADMIN: HOD Matrix
   ══════════════════════════════════════════════════════════ */
function HodsPage() {
  const toast = useAppToast();
  const [list, setList] = useState<any[]>([]);
  const [emps, setEmps] = useState<Employee[]>([]);
  const [search, setSearch] = useState("");
  const [loc, setLoc] = useState("");
  const [dept, setDept] = useState("");
  const [hodEmail, setHodEmail] = useState("");
  const [editId, setEditId] = useState<number | null>(null);
  const [combos, setCombos] = useState<
    { Location: string; Department: string }[]
  >([]);
  const [cLoc, setCLoc] = useState("");
  const [cDept, setCDept] = useState("");

  const load = () =>
    api<any[]>("/api/hods")
      .then(setList)
      .catch(() => setList([]));
  useEffect(() => {
    load();
  }, []);

  useEffect(() => {
    if (search.length >= 2)
      api<Employee[]>(`/api/employees/search?q=${encodeURIComponent(search)}`)
        .then(setEmps)
        .catch(() => setEmps([]));
  }, [search]);

  useEffect(() => {
    const p = new URLSearchParams();
    if (cLoc) p.set("location", cLoc);
    if (cDept) p.set("department", cDept);
    api<{ combinations: any[] }>(`/api/hods/combinations?${p}`)
      .then((r) => setCombos(r.combinations || []))
      .catch(() => setCombos([]));
  }, [cLoc, cDept]);

  const save = async () => {
    const method = editId ? "PUT" : "POST";
    const url = editId ? `/api/hods/${editId}` : "/api/hods";
    await api(url, {
      method,
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        location: loc,
        department: dept,
        hodEmail,
        hodName: "",
      }),
    });
    toast({ title: editId ? "Updated" : "Created", variant: "success" });
    setLoc("");
    setDept("");
    setHodEmail("");
    setEditId(null);
    load();
  };

  return (
    <div className="space-y-6 animate-fade-in">
      <div>
        <h2 className="font-display text-3xl">HOD Matrix</h2>
        <p className="text-sm text-muted-foreground mt-1">
          Map Location + Department combinations to HOD approvers.
        </p>
      </div>

      <div className="grid xl:grid-cols-3 gap-6">
        <Card className="xl:col-span-2 glass" data-tour="hod-combinations">
          <CardContent className="p-6 space-y-4">
            <div className="grid md:grid-cols-2 gap-3">
              <Input
                placeholder="Filter by location"
                value={cLoc}
                onChange={(e) => setCLoc(e.target.value)}
              />
              <Input
                placeholder="Filter by department"
                value={cDept}
                onChange={(e) => setCDept(e.target.value)}
              />
            </div>
            <div className="max-h-48 overflow-auto border rounded-lg divide-y text-sm">
              {combos.map((c, i) => (
                <button
                  key={`${c.Location}-${c.Department}-${i}`}
                  type="button"
                  className="w-full text-left p-2.5 hover:bg-muted/50 transition-colors"
                  onClick={() => {
                    setLoc(c.Location);
                    setDept(c.Department);
                  }}
                >
                  <span className="font-medium">{c.Location}</span> /{" "}
                  {c.Department}
                </button>
              ))}
            </div>
            <div className="grid md:grid-cols-2 gap-3">
              <Input
                placeholder="Location"
                value={loc}
                onChange={(e) => setLoc(e.target.value)}
              />
              <Input
                placeholder="Department"
                value={dept}
                onChange={(e) => setDept(e.target.value)}
              />
            </div>
            <Input
              placeholder="Search employee for HOD..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
            <div className="max-h-28 overflow-auto border rounded-lg divide-y text-sm">
              {emps.map((e) => (
                <button
                  key={e.EmpEmail}
                  type="button"
                  className="w-full text-left p-2 hover:bg-muted/50 transition-colors"
                  onClick={() => setHodEmail(e.EmpEmail)}
                >
                  {e.EmpName} ({e.EmpEmail})
                </button>
              ))}
            </div>
            <Input
              placeholder="HOD Email"
              value={hodEmail}
              onChange={(e) => setHodEmail(e.target.value)}
            />
            <div className="flex gap-2">
              <Button onClick={save}>{editId ? "Update" : "Save"}</Button>
              {editId && (
                <Button
                  variant="outline"
                  onClick={() => {
                    setEditId(null);
                    setLoc("");
                    setDept("");
                    setHodEmail("");
                  }}
                >
                  Cancel
                </Button>
              )}
            </div>
          </CardContent>
        </Card>

        <Card className="glass">
          <CardHeader className="pb-3">
            <CardTitle className="text-base">Configured HODs</CardTitle>
          </CardHeader>
          <CardContent className="max-h-[500px] overflow-auto space-y-2">
            {list.map((r) => (
              <div
                key={r.Id}
                className="p-3 border rounded-lg text-sm hover:bg-muted/20 transition-colors"
              >
                <div className="font-medium">
                  {r.Location} / {r.Department}
                </div>
                <div className="text-xs text-muted-foreground">
                  {r.HodEmail}
                </div>
                <div className="flex items-center gap-2 mt-2">
                  <Badge
                    variant={r.Active ? "success" : "secondary"}
                    className="text-[10px]"
                  >
                    {r.Active ? "Active" : "Off"}
                  </Badge>
                  <Button
                    size="sm"
                    variant="outline"
                    className="h-6 text-[10px]"
                    onClick={() => {
                      setEditId(r.Id);
                      setLoc(r.Location);
                      setDept(r.Department);
                      setHodEmail(r.HodEmail);
                    }}
                  >
                    Edit
                  </Button>
                  <Button
                    size="sm"
                    variant="ghost"
                    className="h-6 text-[10px] text-destructive"
                    onClick={async () => {
                      await api(`/api/hods/${r.Id}`, { method: "DELETE" });
                      toast({ title: "Deactivated", variant: "success" });
                      load();
                    }}
                  >
                    Remove
                  </Button>
                </div>
              </div>
            ))}
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   ADMIN: Analytics
   ══════════════════════════════════════════════════════════ */
const CHART_COLORS = [
  "#0c4a6e",
  "#059669",
  "#d97706",
  "#dc2626",
  "#7c3aed",
  "#0891b2",
];

function AnalyticsPage() {
  const [ov, setOv] = useState<any>(null);
  const [byDept, setByDept] = useState<any[]>([]);
  const [byLoc, setByLoc] = useState<any[]>([]);
  const [byUser, setByUser] = useState<any[]>([]);

  useEffect(() => {
    Promise.all([
      api("/api/analytics/overview"),
      api<any[]>("/api/analytics/by-department"),
      api<any[]>("/api/analytics/by-location"),
      api<any[]>("/api/analytics/by-user"),
    ])
      .then(([o, d, l, u]) => {
        setOv(o);
        setByDept(d);
        setByLoc(l);
        setByUser(u);
      })
      .catch(() => undefined);
  }, []);

  const pie = useMemo(
    () =>
      ov
        ? [
            { name: "Total", value: ov.totalDocuments || 0 },
            { name: "Controlled", value: ov.controlledDocuments || 0 },
            { name: "Approved", value: ov.approvedDocuments || 0 },
            { name: "Pending", value: ov.pendingApprovals || 0 },
          ]
        : [],
    [ov]
  );

  return (
    <div className="space-y-6 animate-fade-in" data-tour="analytics-main">
      <h2 className="font-display text-3xl">Analytics</h2>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 stagger-children">
        <Stat
          title="Documents"
          value={String(ov?.totalDocuments || 0)}
          icon={<FileArchive className="h-5 w-5" />}
        />
        <Stat
          title="Controlled"
          value={String(ov?.controlledDocuments || 0)}
          icon={<ShieldCheck className="h-5 w-5" />}
          accent="bg-amber-500/10 text-amber-600"
        />
        <Stat
          title="Pending"
          value={String(ov?.pendingApprovals || 0)}
          icon={<Clock3 className="h-5 w-5" />}
          accent="bg-orange-500/10 text-orange-600"
        />
        <Stat
          title="Uploaders"
          value={String(ov?.uniqueUploaders || 0)}
          icon={<Users className="h-5 w-5" />}
          accent="bg-teal-500/10 text-teal-600"
        />
      </div>

      <Card className="glass">
        <CardHeader className="pb-2">
          <CardTitle className="text-base">By Department</CardTitle>
        </CardHeader>
        <CardContent className="h-[280px]">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={byDept}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="Department" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip />
              <Bar
                dataKey="DocumentCount"
                fill="#0c4a6e"
                radius={[6, 6, 0, 0]}
              />
            </BarChart>
          </ResponsiveContainer>
        </CardContent>
      </Card>

      <div className="grid xl:grid-cols-2 gap-6">
        <Card className="glass">
          <CardHeader className="pb-2">
            <CardTitle className="text-base">By Location</CardTitle>
          </CardHeader>
          <CardContent className="h-[260px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={byLoc}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis dataKey="Location" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip />
                <Bar
                  dataKey="DocumentCount"
                  fill="#059669"
                  radius={[6, 6, 0, 0]}
                />
              </BarChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
        <Card className="glass">
          <CardHeader className="pb-2">
            <CardTitle className="text-base">Document Mix</CardTitle>
          </CardHeader>
          <CardContent className="h-[260px]">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={pie}
                  dataKey="value"
                  nameKey="name"
                  outerRadius={90}
                  label={{ fontSize: 11 }}
                >
                  {pie.map((_, i) => (
                    <Cell
                      key={i}
                      fill={CHART_COLORS[i % CHART_COLORS.length]}
                    />
                  ))}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
      </div>

      <Card className="glass">
        <CardHeader className="pb-2">
          <CardTitle className="text-base">Top Uploaders</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="max-h-48 overflow-auto divide-y rounded-lg border">
            {byUser.map((u, i) => (
              <div
                key={u.CreatorEmail}
                className="px-4 py-2.5 text-sm flex items-center justify-between hover:bg-muted/20 transition-colors"
              >
                <span className="flex items-center gap-2">
                  <span className="text-xs text-muted-foreground w-5">
                    {i + 1}.
                  </span>
                  {u.CreatorEmail}
                </span>
                <Badge variant="outline" className="text-[10px]">
                  {u.DocumentCount}
                </Badge>
              </div>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   ADMIN: Audit Log
   ══════════════════════════════════════════════════════════ */
function AuditPage() {
  const [logs, setLogs] = useState<any[]>([]);
  const [action, setAction] = useState("");
  const [user, setUser] = useState("");

  const load = () => {
    const p = new URLSearchParams();
    if (action) p.set("action", action);
    if (user) p.set("user", user);
    p.set("pageSize", "100");
    api<{ logs: any[] }>(`/api/audit-log?${p}`)
      .then((r) => setLogs(r.logs || []))
      .catch(() => setLogs([]));
  };
  useEffect(() => {
    load();
  }, []);

  return (
    <div className="space-y-6 animate-fade-in">
      <h2 className="font-display text-3xl">Audit Trail</h2>

      <Card className="glass">
        <CardContent className="p-5 space-y-4">
          <div className="flex flex-wrap gap-3" data-tour="audit-filters">
            <Input
              placeholder="Action"
              value={action}
              onChange={(e) => setAction(e.target.value)}
              className="w-48"
            />
            <Input
              placeholder="User email"
              value={user}
              onChange={(e) => setUser(e.target.value)}
              className="w-56"
            />
            <Button onClick={load}>Filter</Button>
          </div>
          <div className="max-h-[600px] overflow-auto rounded-lg border divide-y">
            {logs.map((l) => (
              <div
                key={l.Id}
                className="p-4 hover:bg-muted/20 transition-colors text-sm space-y-1.5"
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="font-medium text-xs">{l.Action}</span>
                  <span className="text-[10px] text-muted-foreground">
                    {new Date(l.CreatedAt).toLocaleString()}
                  </span>
                </div>
                <p className="text-xs text-muted-foreground">
                  {l.UserEmail || "System"}
                  {l.Reason ? ` — ${l.Reason}` : ""}
                </p>
                <div className="grid md:grid-cols-2 gap-2 text-xs">
                  <pre className="p-2 bg-muted/40 rounded-md overflow-auto max-h-20 whitespace-pre-wrap">
                    Before: {l.BeforeState || "—"}
                  </pre>
                  <pre className="p-2 bg-muted/40 rounded-md overflow-auto max-h-20 whitespace-pre-wrap">
                    After: {l.AfterState || "—"}
                  </pre>
                </div>
              </div>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   PUBLIC VIEWER
   ══════════════════════════════════════════════════════════ */
function PublicViewerPage() {
  const { token = "" } = useParams();
  const [meta, setMeta] = useState<any>(null);
  const [error, setError] = useState("");

  useEffect(() => {
    api(`/api/public/meta/${token}`)
      .then(setMeta)
      .catch((e) => setError(String(e.message || "Invalid link")));
  }, [token]);

  return (
    <div className="min-h-screen bg-page-gradient no-print-view p-4 md:p-8">
      <div className="max-w-5xl mx-auto space-y-4">
        <div className="flex items-center gap-3">
          <img src="/l.png" alt="Logo" className="h-8 w-8" />
          <h1 className="font-display text-3xl text-primary">
            Public Document View
          </h1>
        </div>
        {error && (
          <p className="text-destructive p-4 rounded-lg border bg-red-50">
            {error}
          </p>
        )}
        {meta && (
          <div className="flex gap-3 flex-wrap text-sm">
            <Badge variant="outline">{meta.Title}</Badge>
            <Badge variant="secondary">v{meta.CurrentVersion}</Badge>
            {meta.IsControlled && <Badge variant="warning">Controlled</Badge>}
          </div>
        )}
        <iframe
          title="public-doc"
          className="w-full h-[82vh] rounded-xl border bg-white shadow-lg"
          src={`/api/public/view/${token}`}
        />
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   Admin Gate
   ══════════════════════════════════════════════════════════ */
function AdminGate({ children }: { children: React.ReactNode }) {
  const { user } = useAuth();
  if (!user?.isAdmin)
    return (
      <div className="max-w-md mx-auto py-20 text-center space-y-4">
        <Lock className="h-12 w-12 text-muted-foreground/40 mx-auto" />
        <p className="text-muted-foreground">
          This section is restricted to administrators.
        </p>
      </div>
    );
  return <>{children}</>;
}

/* ══════════════════════════════════════════════════════════
   Main App Router
   ══════════════════════════════════════════════════════════ */
function AuthenticatedApp() {
  const { loading, error, user } = useAuth();
  if (loading || error || !user) return <LoadingScreen />;

  return (
    <ProtectedLayout>
      <Routes>
        <Route path="/" element={<DashboardPage />} />
        <Route path="/upload" element={<UploadPage />} />
        <Route path="/groups" element={<GroupsPage />} />
        <Route path="/approvals" element={<ApprovalsPage />} />
        <Route path="/verify" element={<VerifyPage />} />
        <Route
          path="/admin/users"
          element={
            <AdminGate>
              <AdminUsersPage />
            </AdminGate>
          }
        />
        <Route
          path="/admin/hods"
          element={
            <AdminGate>
              <HodsPage />
            </AdminGate>
          }
        />
        <Route
          path="/admin/analytics"
          element={
            <AdminGate>
              <AnalyticsPage />
            </AdminGate>
          }
        />
        <Route
          path="/admin/audit"
          element={
            <AdminGate>
              <AuditPage />
            </AdminGate>
          }
        />
        <Route
          path="*"
          element={
            <div className="max-w-md mx-auto py-20 text-center space-y-4">
              <Sparkles className="h-8 w-8 text-muted-foreground/40 mx-auto" />
              <p className="text-muted-foreground">
                Page not found.{" "}
                <Link to="/" className="text-primary underline">
                  Go to Dashboard
                </Link>
              </p>
            </div>
          }
        />
      </Routes>
    </ProtectedLayout>
  );
}

export default function App() {
  return (
    <ToastProvider>
      <BrowserRouter>
        <Routes>
          <Route path="/public/:token" element={<PublicViewerPage />} />
          <Route
            path="*"
            element={
              <AuthProvider>
                <AuthenticatedApp />
              </AuthProvider>
            }
          />
        </Routes>
      </BrowserRouter>
    </ToastProvider>
  );
}
