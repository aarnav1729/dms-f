import React, { useEffect, useMemo, useState } from "react";
import { BrowserRouter, Link, NavLink, Route, Routes, useLocation, useNavigate, useParams } from "react-router-dom";
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
  Clock3,
  FileArchive,
  FileCheck2,
  FileUp,
  LogOut,
  Menu,
  Search,
  ShieldCheck,
  Sparkles,
  UserPlus2,
  Users,
  X,
} from "lucide-react";
import { AuthProvider, useAuth } from "@/contexts/AuthContext";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import { Badge } from "@/components/ui/badge";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

type Doc = {
  Id: number;
  Title: string;
  Description: string;
  FileName: string;
  CurrentVersion: number;
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

const ADMIN_NAV = [
  { to: "/admin/users", label: "Users" },
  { to: "/admin/hods", label: "HOD Matrix" },
  { to: "/admin/analytics", label: "Analytics" },
  { to: "/admin/audit", label: "Audit" },
];

const USER_NAV = [
  { to: "/", label: "Dashboard" },
  { to: "/upload", label: "Upload" },
  { to: "/approvals", label: "Approvals" },
  { to: "/verify", label: "Verify Copy" },
];

const BASE_REDIRECT =
  typeof window !== "undefined"
    ? `${window.location.origin}`
    : "https://dms.premierenergies.com";

async function api<T>(url: string, init?: RequestInit): Promise<T> {
  const response = await fetch(url, { credentials: "include", ...init });
  if (!response.ok) {
    const fallback = await response.text().catch(() => "Request failed");
    throw new Error(fallback || `HTTP ${response.status}`);
  }
  return response.json() as Promise<T>;
}

function MouseGlow() {
  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      const x = `${(e.clientX / window.innerWidth) * 100}%`;
      const y = `${(e.clientY / window.innerHeight) * 100}%`;
      document.documentElement.style.setProperty("--spotlight-x", x);
      document.documentElement.style.setProperty("--spotlight-y", y);
    };
    window.addEventListener("mousemove", onMove);
    return () => window.removeEventListener("mousemove", onMove);
  }, []);
  return <div className="global-spotlight" />;
}

function LoadingScreen() {
  return (
    <div className="min-h-screen bg-page-gradient flex items-center justify-center p-6">
      <Card className="glass max-w-xl w-full animate-fade-in">
        <CardHeader>
          <CardTitle className="font-display text-4xl tracking-wide">Premier DMS</CardTitle>
          <CardDescription>Initializing your secure workspace...</CardDescription>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="h-2 rounded-full bg-muted overflow-hidden">
            <div className="h-full w-1/2 bg-primary animate-shimmer bg-[length:200%_100%]" />
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

function TutorialAssistant() {
  const location = useLocation();
  const navigate = useNavigate();
  const [open, setOpen] = useState(() => localStorage.getItem("dms_tutorial_open") === "1");

  useEffect(() => {
    localStorage.setItem("dms_tutorial_open", open ? "1" : "0");
  }, [open]);

  const tips = useMemo(() => {
    if (location.pathname.startsWith("/upload")) {
      return [
        "Choose Controlled only when workflow approvals are mandatory.",
        "For controlled docs, enter a revision reason before upload.",
        "Use Custom Group for targeted sharing; Department and Company are broadcast scopes.",
      ];
    }
    if (location.pathname.startsWith("/admin/hods")) {
      return [
        "Start from Location + Department combinations list to avoid typo mismatches.",
        "Pick HOD from employee search to keep the email valid.",
        "Any unmapped combo skips HOD and flags Document Controller review.",
      ];
    }
    if (location.pathname.startsWith("/admin/users")) {
      return [
        "Create users with minimum required fields first, then refine via edit.",
        "Use search before creating to avoid duplicates.",
        "Delete operation is soft-delete (ActiveFlag=0).",
      ];
    }
    return [
      "Use quick search with controlled filter to find documents fast.",
      "Open a document to see timeline, links, and inline viewer in one flow.",
      "Verify Copy page validates controlled documents by file hash.",
    ];
  }, [location.pathname]);

  return (
    <>
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        className="fixed bottom-20 right-4 z-50 rounded-full px-4 py-2 bg-primary text-primary-foreground shadow-xl hover:scale-105 transition"
      >
        <BookOpen className="h-4 w-4 inline mr-2" />Tutorial
      </button>
      <AnimatePresence>
        {open ? (
          <motion.div
            initial={{ opacity: 0, y: 18 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: 12 }}
            className="fixed bottom-36 right-4 z-50 w-[min(90vw,420px)]"
          >
            <Card className="glass shadow-2xl">
              <CardHeader className="pb-3">
                <CardTitle className="text-lg">Guided Mode</CardTitle>
                <CardDescription>Suggestions based on your current page</CardDescription>
              </CardHeader>
              <CardContent className="space-y-2 text-sm">
                {tips.map((tip, i) => (
                  <div key={i} className="p-2 rounded-md bg-muted/50 border">{tip}</div>
                ))}
                <div className="flex gap-2 pt-2 flex-wrap">
                  <Button size="sm" variant="outline" onClick={() => navigate("/upload")}>Go to Upload</Button>
                  <Button size="sm" variant="outline" onClick={() => navigate("/verify")}>Verify Copy</Button>
                  <Button size="sm" onClick={() => setOpen(false)}>Close</Button>
                </div>
              </CardContent>
            </Card>
          </motion.div>
        ) : null}
      </AnimatePresence>
    </>
  );
}

type TourStep = {
  route: string;
  selector: string;
  title: string;
  body: string;
};

function FirstTimeTourOverlay({ isAdmin }: { isAdmin: boolean }) {
  const navigate = useNavigate();
  const location = useLocation();
  const [index, setIndex] = useState(0);
  const [enabled, setEnabled] = useState(false);
  const [rect, setRect] = useState<DOMRect | null>(null);

  const steps = useMemo<TourStep[]>(
    () => [
      { route: "/", selector: "[data-tour='dashboard-search']", title: "Dashboard Search", body: "Find documents instantly by content, metadata, department, location, and control type." },
      { route: "/", selector: "[data-tour='dashboard-doc-card']", title: "Document Cards", body: "Open, view inline, download, and inspect complete history and public links from each card." },
      { route: "/upload", selector: "[data-tour='upload-controlled']", title: "Controlled Workflow", body: "Enable this when document controller approval and strict versioning are mandatory." },
      { route: "/upload", selector: "[data-tour='upload-submit']", title: "Upload Action", body: "Upload single files or folders. Controlled uploads should include a mandatory revision reason." },
      { route: "/approvals", selector: "[data-tour='approvals-list']", title: "Approvals Queue", body: "Approve or reject pending items by workflow stage. HOD-skipped items are visibly flagged." },
      { route: "/verify", selector: "[data-tour='verify-panel']", title: "Copy Verification", body: "Anyone can upload a file to verify whether it matches an official controlled copy." },
      ...(isAdmin
        ? [
            { route: "/admin/users", selector: "[data-tour='users-create']", title: "User CRUD", body: "Create, update, and deactivate users from the EMP-backed master. Admin management is separate." },
            { route: "/admin/hods", selector: "[data-tour='hod-combinations']", title: "HOD Combinations", body: "Search and pick exact Location + Department combinations before assigning HOD." },
            { route: "/admin/analytics", selector: "[data-tour='analytics-main']", title: "Analytics", body: "Track usage and drilldowns by department, location, and uploader behavior." },
            { route: "/admin/audit", selector: "[data-tour='audit-filters']", title: "Audit Trail", body: "Filter immutable action logs with before/after states, actor, reason, and timestamp." },
          ]
        : []),
    ],
    [isAdmin]
  );

  useEffect(() => {
    if (localStorage.getItem("dms_tour_done_v1") === "1") return;
    const t = window.setTimeout(() => setEnabled(true), 450);
    return () => window.clearTimeout(t);
  }, []);

  useEffect(() => {
    if (!enabled) return;
    const step = steps[index];
    if (!step) return;
    if (location.pathname !== step.route) {
      navigate(step.route);
      return;
    }

    let tries = 0;
    const maxTries = 50;
    const timer = window.setInterval(() => {
      const el = document.querySelector(step.selector) as HTMLElement | null;
      if (el) {
        el.scrollIntoView({ behavior: "smooth", block: "center" });
        const r = el.getBoundingClientRect();
        setRect(r);
        window.clearInterval(timer);
      } else if (tries++ > maxTries) {
        setRect(null);
        window.clearInterval(timer);
      }
    }, 80);
    return () => window.clearInterval(timer);
  }, [enabled, index, location.pathname, navigate, steps]);

  if (!enabled || !steps[index]) return null;

  const step = steps[index];
  const done = () => {
    localStorage.setItem("dms_tour_done_v1", "1");
    setEnabled(false);
  };

  return (
    <div className="fixed inset-0 z-[70] pointer-events-none">
      <div className="absolute inset-0 bg-slate-950/70" />
      {rect ? (
        <div
          className="absolute rounded-xl border-2 border-cyan-300/90 shadow-[0_0_0_9999px_rgba(2,6,23,0.72)] transition-all duration-300"
          style={{
            left: rect.left - 8,
            top: rect.top - 8,
            width: rect.width + 16,
            height: rect.height + 16,
          }}
        />
      ) : null}
      <div
        className="absolute pointer-events-auto w-[min(92vw,440px)]"
        style={{
          left: rect ? Math.max(12, Math.min(window.innerWidth - 452, rect.left)) : 12,
          top: rect ? Math.min(window.innerHeight - 240, rect.bottom + 14) : 80,
        }}
      >
        <Card className="glass shadow-2xl">
          <CardHeader className="pb-2">
            <CardTitle className="text-lg">{step.title}</CardTitle>
            <CardDescription>Step {index + 1} of {steps.length}</CardDescription>
          </CardHeader>
          <CardContent className="space-y-3 text-sm">
            <p>{step.body}</p>
            <div className="flex gap-2 flex-wrap">
              <Button size="sm" variant="outline" disabled={index === 0} onClick={() => setIndex((v) => Math.max(0, v - 1))}>Back</Button>
              <Button size="sm" onClick={() => (index >= steps.length - 1 ? done() : setIndex((v) => v + 1))}>
                {index >= steps.length - 1 ? "Finish Tour" : "Next"}
              </Button>
              <Button size="sm" variant="ghost" onClick={done}>Skip</Button>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

function ProtectedLayout({ children }: { children: React.ReactNode }) {
  const { user, logout } = useAuth();
  const navigate = useNavigate();
  const [mobileOpen, setMobileOpen] = useState(false);

  const navItems = [...USER_NAV, ...(user?.isAdmin ? ADMIN_NAV : [])];

  return (
    <div className="min-h-screen bg-page-gradient relative overflow-x-hidden flex flex-col">
      <MouseGlow />
      <div className="grid-pattern fixed inset-0 pointer-events-none" />

      <header className="sticky top-0 z-40 backdrop-blur-xl border-b border-border/60 bg-background/70">
        <div className="w-full px-3 md:px-8 py-3 flex items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="h-11 w-11 rounded-2xl bg-gradient-to-br from-sky-500 to-emerald-500 shadow-lg shadow-sky-400/30 animate-float-gentle" />
            <div>
              <h1 className="font-display text-3xl tracking-wide leading-none">DMS</h1>
              <p className="text-xs text-muted-foreground -mt-1">Premier Energies Document Management</p>
            </div>
          </div>

          <div className="hidden md:flex items-center gap-2 text-xs md:text-sm">
            <Badge variant="info">{user?.department || "Department"}</Badge>
            <Badge variant="secondary">{user?.location || "Location"}</Badge>
            <Button variant="ghost" size="icon" onClick={logout} title="Logout">
              <LogOut className="h-4 w-4" />
            </Button>
          </div>

          <div className="md:hidden flex items-center gap-2">
            <Button variant="ghost" size="icon" onClick={() => setMobileOpen(true)} title="Menu">
              <Menu className="h-5 w-5" />
            </Button>
          </div>
        </div>
      </header>

      <AnimatePresence>
        {mobileOpen ? (
          <motion.div
            initial={{ x: "100%", opacity: 0.7 }}
            animate={{ x: 0, opacity: 1 }}
            exit={{ x: "100%", opacity: 0.9 }}
            transition={{ duration: 0.45, ease: [0.2, 0.8, 0.2, 1] }}
            className="fixed inset-0 z-[120] text-white"
          >
            <div className="absolute inset-0 bg-gradient-to-b from-[#004b8d] via-[#007f6d] to-[#f08a24]" />
            <div className="absolute inset-0 bg-[radial-gradient(circle_at_20%_15%,rgba(255,255,255,0.25),transparent_34%),radial-gradient(circle_at_85%_80%,rgba(255,255,255,0.18),transparent_38%)]" />
            <div className="relative h-full w-full p-6 flex flex-col text-white">
              <div className="flex items-center justify-between mb-8">
                <div>
                  <h2 className="font-display text-5xl tracking-wide drop-shadow-lg">Menu</h2>
                  <p className="text-white/90 text-sm">Navigate your DMS workspace</p>
                </div>
                <Button variant="ghost" size="icon" className="!text-white hover:!bg-white/15 transition-all duration-300 hover:rotate-90" onClick={() => setMobileOpen(false)}>
                  <X className="h-6 w-6" />
                </Button>
              </div>

              <div className="flex-1 space-y-2 overflow-auto">
                {navItems.map((item, idx) => (
                  <motion.div
                    key={item.to}
                    initial={{ opacity: 0, y: 16 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: idx * 0.06 }}
                  >
                    <NavLink
                      to={item.to}
                      onClick={() => setMobileOpen(false)}
                      className="block rounded-2xl border border-white/45 px-5 py-4 text-xl bg-white/15 text-white shadow-lg backdrop-blur-md hover:bg-white/25 hover:translate-x-1 hover:scale-[1.01] transition-all duration-300"
                    >
                      {item.label}
                    </NavLink>
                  </motion.div>
                ))}
              </div>

              <div className="pt-4 border-t border-white/35 space-y-3">
                <div className="text-sm text-white">{user?.email}</div>
                <div className="flex gap-2">
                  <Button className="flex-1 !bg-white !text-[#004b8d] hover:!bg-white/90 hover:scale-[1.02] transition-all" onClick={() => { setMobileOpen(false); navigate("/upload"); }}>New Document</Button>
                  <Button className="flex-1 !text-white !border !border-white/60 hover:!bg-white/15 hover:scale-[1.02] transition-all" variant="ghost" onClick={logout}>Logout</Button>
                </div>
              </div>
            </div>
          </motion.div>
        ) : null}
      </AnimatePresence>

      <main className="w-full px-3 md:px-8 py-6 relative z-10 flex-1">
        <div className="mb-6 hidden md:flex flex-wrap gap-2">
          {navItems.map((item) => (
            <NavLink key={item.to} to={item.to} className="px-4 py-2 rounded-full border border-border hover:bg-card transition">
              {item.label}
            </NavLink>
          ))}
          <Button variant="outline" onClick={() => navigate("/upload")}>New Document</Button>
        </div>
        {children}
      </main>

      <footer className="mt-auto border-t border-border/60 bg-background/85 backdrop-blur-xl sticky bottom-0">
        <div className="w-full px-3 md:px-8 py-3 text-xs text-muted-foreground flex items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <img src="/l.png" alt="Premier Energies logo" className="h-7 w-auto" />
            <span>Secure DMS Platform</span>
          </div>
          <span>© {new Date().getFullYear()} Premier Energies. All Rights Reserved.</span>
        </div>
      </footer>

      <TutorialAssistant />
      <FirstTimeTourOverlay isAdmin={Boolean(user?.isAdmin)} />
    </div>
  );
}

function DashboardPage() {
  const { user } = useAuth();
  const [documents, setDocuments] = useState<Doc[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(false);
  const [query, setQuery] = useState("");
  const [controlled, setControlled] = useState("all");
  const [department, setDepartment] = useState("");
  const [location, setLocation] = useState("");
  const [selected, setSelected] = useState<Doc | null>(null);
  const [history, setHistory] = useState<any[]>([]);

  const fetchDocs = async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (query) params.set("search", query);
      if (department) params.set("department", department);
      if (location) params.set("location", location);
      if (controlled !== "all") params.set("isControlled", String(controlled === "controlled"));
      params.set("pageSize", "40");
      const data = await api<{ documents: Doc[]; total: number }>(`/api/documents?${params.toString()}`);
      setDocuments(data.documents || []);
      setTotal(data.total || 0);
    } catch (error) {
      console.error(error);
      const msg = String((error as Error)?.message || "");
      if (!msg.includes("401") && !msg.toLowerCase().includes("unauth")) {
        console.warn("Document load deferred/failed:", msg);
      }
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (!user?.email) return;
    fetchDocs();
  }, [user?.email]);

  const openDoc = async (doc: Doc) => {
    setSelected(doc);
    try {
      const logs = await api<{ logs: any[] }>(`/api/documents/${doc.Id}/audit`);
      setHistory(logs.logs || []);
    } catch {
      setHistory([]);
    }
  };

  return (
    <div className="space-y-4 animate-fade-in">
      <Card className="glass" data-tour="dashboard-search">
        <CardHeader>
          <CardTitle className="flex items-center gap-2"><Search className="h-5 w-5" />Search & Advanced Filters</CardTitle>
          <CardDescription>Search metadata, content, title, creator, location and document controller flow status.</CardDescription>
        </CardHeader>
        <CardContent className="grid lg:grid-cols-5 gap-3">
          <Input placeholder="Any keyword in title/content/metadata" value={query} onChange={(e) => setQuery(e.target.value)} />
          <Input placeholder="Department" value={department} onChange={(e) => setDepartment(e.target.value)} />
          <Input placeholder="Location" value={location} onChange={(e) => setLocation(e.target.value)} />
          <Select value={controlled} onValueChange={setControlled}>
            <SelectTrigger><SelectValue placeholder="Controlled filter" /></SelectTrigger>
            <SelectContent>
              <SelectItem value="all">All</SelectItem>
              <SelectItem value="controlled">Controlled</SelectItem>
              <SelectItem value="uncontrolled">Not Controlled</SelectItem>
            </SelectContent>
          </Select>
          <Button onClick={fetchDocs}>{loading ? "Searching..." : "Apply"}</Button>
        </CardContent>
      </Card>

      <div className="grid md:grid-cols-4 gap-3">
        <Stat title="Documents" value={String(total)} icon={<FileArchive className="h-5 w-5" />} />
        <Stat title="Controlled" value={String(documents.filter((d) => d.IsControlled).length)} icon={<ShieldCheck className="h-5 w-5" />} />
        <Stat title="Pending" value={String(documents.filter((d) => d.ApprovalStatus?.startsWith("pending")).length)} icon={<Clock3 className="h-5 w-5" />} />
        <Stat title="Approved" value={String(documents.filter((d) => d.Status === "approved").length)} icon={<CheckCircle2 className="h-5 w-5" />} />
      </div>

      <div className="grid xl:grid-cols-2 gap-4">
        {documents.map((doc, index) => (
          <motion.div
            key={doc.Id}
            initial={{ opacity: 0, y: 14 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: index * 0.03 }}
          >
            <Card className="glass hover:shadow-xl hover:-translate-y-1 transition-all duration-300" data-tour={index === 0 ? "dashboard-doc-card" : undefined}>
              <CardContent className="pt-6 space-y-2">
                <div className="flex items-start justify-between gap-3">
                  <div>
                    <h3 className="font-semibold text-lg leading-tight">{doc.Title}</h3>
                    <p className="text-xs text-muted-foreground">{doc.FileName}</p>
                  </div>
                  <Badge variant={doc.IsControlled ? "warning" : "secondary"}>{doc.IsControlled ? "Controlled" : "Open"}</Badge>
                </div>
                <div className="flex flex-wrap gap-2">
                  <Badge variant="outline">v{doc.CurrentVersion}</Badge>
                  <Badge variant="info">{doc.ApprovalStatus || "none"}</Badge>
                  {doc.HodSkipped ? <Badge variant="warning">HOD Skipped</Badge> : null}
                </div>
                <p className="text-sm text-muted-foreground line-clamp-2">{doc.Description || "No description"}</p>
                <div className="flex gap-2 pt-1 flex-wrap">
                  <Button variant="outline" onClick={() => openDoc(doc)}>Open</Button>
                  <a href={`/api/documents/${doc.Id}/view`} target="_blank" rel="noreferrer" className="inline-flex">
                    <Button variant="secondary">Inline View</Button>
                  </a>
                  <a href={`/api/documents/${doc.Id}/download`} target="_blank" rel="noreferrer" className="inline-flex">
                    <Button>Download</Button>
                  </a>
                </div>
              </CardContent>
            </Card>
          </motion.div>
        ))}
      </div>

      {selected ? (
        <Card className="glass">
          <CardHeader>
            <CardTitle>{selected.Title}</CardTitle>
            <CardDescription>Viewer + history timeline + public access links</CardDescription>
          </CardHeader>
          <CardContent className="space-y-3">
            <iframe title="viewer" className="w-full h-[420px] rounded-xl border" src={`/api/documents/${selected.Id}/view`} />
            <PublicLinksPanel docId={selected.Id} />
            <div>
              <h4 className="font-semibold mb-2">History</h4>
              <div className="max-h-52 overflow-auto border rounded-md">
                {history.length === 0 ? <p className="p-3 text-sm text-muted-foreground">No history found.</p> : null}
                {history.map((h, i) => (
                  <div key={i} className="p-3 border-b last:border-b-0 text-sm">
                    <div className="font-medium">{h.Action}</div>
                    <div className="text-muted-foreground">{h.UserEmail || "System"} · {new Date(h.CreatedAt).toLocaleString()}</div>
                    {h.Reason ? <div className="text-xs mt-1">Reason: {h.Reason}</div> : null}
                  </div>
                ))}
              </div>
            </div>
          </CardContent>
        </Card>
      ) : null}
    </div>
  );
}

function PublicLinksPanel({ docId }: { docId: number }) {
  const [links, setLinks] = useState<any[]>([]);
  const [expiresAt, setExpiresAt] = useState("");

  const load = async () => {
    try {
      const data = await api<{ links: any[] }>(`/api/documents/${docId}/public-links`);
      setLinks(data.links || []);
    } catch {
      setLinks([]);
    }
  };

  useEffect(() => {
    load();
  }, [docId]);

  const createLink = async () => {
    try {
      await api("/api/public-links", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ docId, expiresAt: expiresAt || null }),
      });
      load();
    } catch {
      alert("Failed to create public link");
    }
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-lg">Public View Links</CardTitle>
      </CardHeader>
      <CardContent className="space-y-2">
        <div className="flex flex-wrap gap-2">
          <Input type="datetime-local" value={expiresAt} onChange={(e) => setExpiresAt(e.target.value)} className="max-w-xs" />
          <Button onClick={createLink}>Generate Link</Button>
        </div>
        <div className="space-y-2">
          {links.map((link) => {
            const url = `${BASE_REDIRECT}/public/${link.LinkToken}`;
            return (
              <div key={link.Id} className="p-2 border rounded text-sm flex flex-wrap items-center gap-2 justify-between">
                <a href={url} className="text-primary underline break-all" target="_blank" rel="noreferrer">{url}</a>
                <Button size="sm" variant="outline" onClick={() => navigator.clipboard.writeText(url)}>Copy</Button>
              </div>
            );
          })}
        </div>
      </CardContent>
    </Card>
  );
}

function UploadPage() {
  const [title, setTitle] = useState("");
  const [description, setDescription] = useState("");
  const [metadata, setMetadata] = useState("{}");
  const [isControlled, setIsControlled] = useState(false);
  const [shareScope, setShareScope] = useState("private");
  const [shareGroupId, setShareGroupId] = useState("");
  const [groups, setGroups] = useState<any[]>([]);
  const [reason, setReason] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [folderFiles, setFolderFiles] = useState<File[]>([]);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    api<any[]>("/api/share-groups").then(setGroups).catch(() => setGroups([]));
  }, []);

  const doUpload = async (docFile: File, overrideTitle?: string) => {
    const form = new FormData();
    form.append("file", docFile);
    form.append("title", overrideTitle || title || docFile.name);
    form.append("description", description);
    form.append("metadata", metadata);
    form.append("isControlled", String(isControlled));
    form.append("shareScope", shareScope);
    if (shareScope === "group" && shareGroupId) form.append("shareGroupId", shareGroupId);
    if (reason) form.append("reason", reason);
    await api("/api/documents/upload", { method: "POST", body: form });
  };

  const submit = async () => {
    if (!file && folderFiles.length === 0) {
      alert("Choose a file or folder.");
      return;
    }
    if (isControlled && !reason.trim()) {
      alert("Reason is mandatory for controlled copy/revision uploads.");
      return;
    }
    setSaving(true);
    try {
      if (file) await doUpload(file);
      for (const f of folderFiles) {
        await doUpload(f, f.name);
      }
      setTitle("");
      setDescription("");
      setReason("");
      setFile(null);
      setFolderFiles([]);
      alert("Uploaded successfully.");
    } catch (err) {
      console.error(err);
      alert("Upload failed.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="grid xl:grid-cols-3 gap-4 animate-fade-in">
      <Card className="xl:col-span-2 glass">
        <CardHeader>
          <CardTitle className="flex items-center gap-2"><FileUp className="h-5 w-5" />Document Upload</CardTitle>
          <CardDescription>Upload single documents or entire folders with controlled/non-controlled workflow logic.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="grid md:grid-cols-2 gap-3">
            <div>
              <Label>Title</Label>
              <Input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Document title" />
            </div>
            <div>
              <Label>Revision Reason</Label>
              <Input value={reason} onChange={(e) => setReason(e.target.value)} placeholder="Mandatory for controlled docs" />
            </div>
          </div>

          <div>
            <Label>Description</Label>
            <textarea value={description} onChange={(e) => setDescription(e.target.value)} className="w-full h-24 rounded-md border bg-background p-3 text-sm" />
          </div>

          <div>
            <Label>Metadata JSON (indexed)</Label>
            <textarea value={metadata} onChange={(e) => setMetadata(e.target.value)} className="w-full h-24 rounded-md border bg-background p-3 text-sm font-mono" />
          </div>

          <div className="flex items-center gap-2" data-tour="upload-controlled">
            <Checkbox checked={isControlled} onCheckedChange={(v) => setIsControlled(Boolean(v))} id="controlled" />
            <Label htmlFor="controlled">Requires Document Controller Approval</Label>
          </div>

          <div className="grid md:grid-cols-2 gap-3">
            <div>
              <Label>Share Scope</Label>
              <Select value={shareScope} onValueChange={setShareScope}>
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="private">Private</SelectItem>
                  <SelectItem value="group">Custom Group</SelectItem>
                  <SelectItem value="department">Department</SelectItem>
                  <SelectItem value="company">Company</SelectItem>
                </SelectContent>
              </Select>
            </div>
            {shareScope === "group" ? (
              <div>
                <Label>Share Group</Label>
                <Select value={shareGroupId} onValueChange={setShareGroupId}>
                  <SelectTrigger><SelectValue placeholder="Choose group" /></SelectTrigger>
                  <SelectContent>
                    {groups.map((g) => (
                      <SelectItem value={String(g.Id)} key={g.Id}>{g.Name}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            ) : null}
          </div>

          <div className="grid md:grid-cols-2 gap-3">
            <div>
              <Label>Single File</Label>
              <Input type="file" onChange={(e) => setFile(e.target.files?.[0] || null)} />
            </div>
            <div>
              <Label>Folder Upload</Label>
              <Input
                type="file"
                multiple
                // @ts-expect-error non-standard but supported by chromium
                webkitdirectory="true"
                onChange={(e) => setFolderFiles(Array.from(e.target.files || []))}
              />
            </div>
          </div>

          <div className="flex gap-2 flex-wrap">
            <Button onClick={submit} disabled={saving} data-tour="upload-submit">{saving ? "Uploading..." : "Upload"}</Button>
            <CreateGroupInline onSaved={() => api<any[]>("/api/share-groups").then(setGroups).catch(() => setGroups([]))} />
          </div>
        </CardContent>
      </Card>

      <Card className="glass">
        <CardHeader>
          <CardTitle>Workflow Rules</CardTitle>
        </CardHeader>
        <CardContent className="space-y-2 text-sm">
          <p><Badge variant="warning">Controlled</Badge> Reporting Manager {"->"} HOD {"->"} Document Controller</p>
          <p>HOD auto-skip if matrix not configured; Document Controller gets skip flag.</p>
          <p>Each new controlled copy creates a strict version trail and hash validation.</p>
          <p>All actions are logged in immutable audit history with before/after state.</p>
        </CardContent>
      </Card>
    </div>
  );
}

function CreateGroupInline({ onSaved }: { onSaved: () => void }) {
  const [name, setName] = useState("");
  const [search, setSearch] = useState("");
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [members, setMembers] = useState<string[]>([]);

  useEffect(() => {
    if (search.length < 2) {
      setEmployees([]);
      return;
    }
    api<Employee[]>(`/api/employees/search?q=${encodeURIComponent(search)}`).then(setEmployees).catch(() => setEmployees([]));
  }, [search]);

  const save = async () => {
    if (!name || members.length === 0) return;
    await api("/api/share-groups", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ name, members }),
    });
    setName("");
    setMembers([]);
    setSearch("");
    onSaved();
  };

  return (
    <Card className="w-full md:w-auto">
      <CardContent className="pt-4 space-y-2">
        <Label>Create Custom Group</Label>
        <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="Group name" />
        <Input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search employee" />
        <div className="max-h-24 overflow-auto text-xs border rounded-md">
          {employees.map((e) => (
            <button
              key={e.EmpEmail}
              className="w-full text-left p-2 border-b hover:bg-muted"
              onClick={() => setMembers((prev) => Array.from(new Set([...prev, e.EmpEmail])))}
              type="button"
            >
              {e.EmpName} ({e.EmpEmail})
            </button>
          ))}
        </div>
        <div className="text-xs text-muted-foreground">Selected: {members.join(", ") || "none"}</div>
        <Button variant="outline" size="sm" onClick={save}>Save Group</Button>
      </CardContent>
    </Card>
  );
}

function ApprovalsPage() {
  const [pending, setPending] = useState<Approval[]>([]);
  const [comments, setComments] = useState<Record<number, string>>({});

  const load = async () => {
    try {
      const data = await api<Approval[]>("/api/approvals/pending");
      setPending(data || []);
    } catch {
      setPending([]);
    }
  };

  useEffect(() => {
    load();
  }, []);

  const decide = async (id: number, action: "approve" | "reject") => {
    const url = `/api/approvals/${id}/${action}`;
    await api(url, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ comments: comments[id] || "" }),
    });
    load();
  };

  return (
    <div className="space-y-3 animate-fade-in" data-tour="approvals-list">
      {pending.length === 0 ? <Card><CardContent className="pt-6">No pending approvals.</CardContent></Card> : null}
      {pending.map((p) => (
        <Card key={p.Id} className="glass">
          <CardContent className="pt-6 space-y-2">
            <div className="flex justify-between gap-3 flex-wrap">
              <div>
                <h3 className="font-semibold">{p.Title} <span className="text-muted-foreground">v{p.CurrentVersion}</span></h3>
                <p className="text-xs">Stage: {p.Stage.replace(/_/g, " ")}</p>
              </div>
              <div className="flex gap-2">
                {p.HodSkipped ? <Badge variant="warning">HOD Skipped - Review Carefully</Badge> : null}
                <Badge variant="info">{p.Department} / {p.Location}</Badge>
              </div>
            </div>
            <Input
              placeholder="Comments"
              value={comments[p.Id] || ""}
              onChange={(e) => setComments((prev) => ({ ...prev, [p.Id]: e.target.value }))}
            />
            <div className="flex gap-2 flex-wrap">
              <Button onClick={() => decide(p.Id, "approve")}>Approve</Button>
              <Button variant="destructive" onClick={() => decide(p.Id, "reject")}>Reject</Button>
              <a href={`/api/documents/${p.DocId}/view`} target="_blank" rel="noreferrer"><Button variant="outline">View</Button></a>
            </div>
          </CardContent>
        </Card>
      ))}
    </div>
  );
}

function VerifyPage() {
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<any>(null);

  const verify = async () => {
    if (!file) return;
    const form = new FormData();
    form.append("file", file);
    const data = await api<any>("/api/documents/verify-upload", { method: "POST", body: form });
    setResult(data);
  };

  return (
    <Card className="glass animate-fade-in" data-tour="verify-panel">
      <CardHeader>
        <CardTitle className="flex items-center gap-2"><FileCheck2 className="h-5 w-5" />Controlled Copy Verification</CardTitle>
        <CardDescription>Accessible to all users. Upload a file and confirm whether it matches a controlled official copy.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-3">
        <Input type="file" onChange={(e) => setFile(e.target.files?.[0] || null)} />
        <Button onClick={verify}>Verify</Button>
        {result ? (
          <div className="p-4 rounded-md border bg-card text-sm space-y-1">
            <p><strong>Matched:</strong> {String(result.matched)}</p>
            <p><strong>Hash:</strong> <span className="font-mono text-xs break-all">{result.uploadHash}</span></p>
            {result.documents?.[0] ? <p><strong>Doc:</strong> {result.documents[0].Title} v{result.documents[0].CurrentVersion}</p> : null}
          </div>
        ) : null}
      </CardContent>
    </Card>
  );
}

function AdminUsersPage() {
  const [users, setUsers] = useState<Employee[]>([]);
  const [admins, setAdmins] = useState<any[]>([]);
  const [search, setSearch] = useState("");
  const [newAdmin, setNewAdmin] = useState("");
  const [editingEmail, setEditingEmail] = useState("");
  const [editDept, setEditDept] = useState("");
  const [editLoc, setEditLoc] = useState("");

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
      api<{ users: Employee[] }>(`/api/admin/users?pageSize=100&search=${encodeURIComponent(search)}`),
      api<any[]>("/api/admin/admins"),
    ]);
    setUsers(u.users || []);
    setAdmins(a || []);
  };

  useEffect(() => {
    load().catch(() => undefined);
  }, [search]);

  const addAdmin = async () => {
    await api("/api/admin/admins", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ email: newAdmin }),
    });
    setNewAdmin("");
    load();
  };

  const createUser = async () => {
    await api("/api/admin/users", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(newUser),
    });
    setNewUser({ EmpID: "", EmpName: "", EmpEmail: "", Department: "", Location: "", ReportingManagerID: "" });
    load();
  };

  const startEdit = (u: Employee) => {
    setEditingEmail(u.EmpEmail);
    setEditDept(u.Department || u.Dept || "");
    setEditLoc(u.Location || u.EmpLocation || "");
  };

  const saveEdit = async () => {
    if (!editingEmail) return;
    await api(`/api/admin/users/${encodeURIComponent(editingEmail)}`, {
      method: "PATCH",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ Department: editDept, Location: editLoc }),
    });
    setEditingEmail("");
    load();
  };

  const deleteUser = async (email: string) => {
    if (!window.confirm(`Deactivate ${email}?`)) return;
    await api(`/api/admin/users/${encodeURIComponent(email)}`, { method: "DELETE" });
    load();
  };

  return (
    <div className="space-y-4 animate-fade-in">
      <Card className="glass" data-tour="users-create">
        <CardHeader>
          <CardTitle className="flex items-center gap-2"><UserPlus2 className="h-5 w-5" />Create User</CardTitle>
        </CardHeader>
        <CardContent className="grid md:grid-cols-3 gap-2">
          <Input placeholder="Emp ID" value={newUser.EmpID} onChange={(e) => setNewUser((p) => ({ ...p, EmpID: e.target.value }))} />
          <Input placeholder="Employee Name" value={newUser.EmpName} onChange={(e) => setNewUser((p) => ({ ...p, EmpName: e.target.value }))} />
          <Input placeholder="Email" value={newUser.EmpEmail} onChange={(e) => setNewUser((p) => ({ ...p, EmpEmail: e.target.value }))} />
          <Input placeholder="Department" value={newUser.Department} onChange={(e) => setNewUser((p) => ({ ...p, Department: e.target.value }))} />
          <Input placeholder="Location" value={newUser.Location} onChange={(e) => setNewUser((p) => ({ ...p, Location: e.target.value }))} />
          <Input placeholder="Reporting Manager ID" value={newUser.ReportingManagerID} onChange={(e) => setNewUser((p) => ({ ...p, ReportingManagerID: e.target.value }))} />
          <div className="md:col-span-3"><Button onClick={createUser}>Create User</Button></div>
        </CardContent>
      </Card>

      <div className="grid xl:grid-cols-2 gap-4">
        <Card className="glass">
          <CardHeader>
            <CardTitle className="flex items-center gap-2"><Users className="h-5 w-5" />User CRUD</CardTitle>
          </CardHeader>
          <CardContent className="space-y-2">
            <Input placeholder="Search by name/email/emp id" value={search} onChange={(e) => setSearch(e.target.value)} />
            <div className="max-h-[420px] overflow-auto border rounded-md">
              {users.map((u) => (
                <div className="p-2 border-b text-sm" key={u.EmpEmail}>
                  <div className="font-medium">{u.EmpName} ({u.EmpID})</div>
                  <div className="text-muted-foreground">{u.EmpEmail} · {u.Department || u.Dept} · {u.Location || u.EmpLocation}</div>
                  <div className="flex gap-2 mt-2">
                    <Button size="sm" variant="outline" onClick={() => startEdit(u)}>Edit</Button>
                    <Button size="sm" variant="destructive" onClick={() => deleteUser(u.EmpEmail)}>Deactivate</Button>
                  </div>
                </div>
              ))}
            </div>
            {editingEmail ? (
              <div className="border rounded-md p-2 space-y-2">
                <p className="text-xs text-muted-foreground">Editing: {editingEmail}</p>
                <Input value={editDept} onChange={(e) => setEditDept(e.target.value)} placeholder="Department" />
                <Input value={editLoc} onChange={(e) => setEditLoc(e.target.value)} placeholder="Location" />
                <div className="flex gap-2">
                  <Button size="sm" onClick={saveEdit}>Save</Button>
                  <Button size="sm" variant="outline" onClick={() => setEditingEmail("")}>Cancel</Button>
                </div>
              </div>
            ) : null}
          </CardContent>
        </Card>

        <Card className="glass">
          <CardHeader>
            <CardTitle>Admin Users</CardTitle>
          </CardHeader>
          <CardContent className="space-y-3">
            <div className="flex gap-2">
              <Input placeholder="admin@premierenergies.com" value={newAdmin} onChange={(e) => setNewAdmin(e.target.value)} />
              <Button onClick={addAdmin}>Add</Button>
            </div>
            <div className="max-h-[420px] overflow-auto border rounded-md">
              {admins.map((a) => (
                <div className="p-2 border-b text-sm flex items-center justify-between" key={a.Email}>
                  <span>{a.Email}</span>
                  <Badge variant={a.Active ? "success" : "secondary"}>{a.Active ? "Active" : "Inactive"}</Badge>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

function HodsPage() {
  const [list, setList] = useState<any[]>([]);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [search, setSearch] = useState("");
  const [location, setLocation] = useState("");
  const [department, setDepartment] = useState("");
  const [hodEmail, setHodEmail] = useState("");
  const [editingId, setEditingId] = useState<number | null>(null);

  const [combos, setCombos] = useState<{ Location: string; Department: string }[]>([]);
  const [comboLocationSearch, setComboLocationSearch] = useState("");
  const [comboDepartmentSearch, setComboDepartmentSearch] = useState("");

  const load = () => api<any[]>("/api/hods").then(setList).catch(() => setList([]));
  useEffect(() => { load(); }, []);

  useEffect(() => {
    if (search.length < 2) return;
    api<Employee[]>(`/api/employees/search?q=${encodeURIComponent(search)}`).then(setEmployees).catch(() => setEmployees([]));
  }, [search]);

  const loadCombos = () => {
    const q = new URLSearchParams();
    if (comboLocationSearch) q.set("location", comboLocationSearch);
    if (comboDepartmentSearch) q.set("department", comboDepartmentSearch);
    api<{ combinations: { Location: string; Department: string }[] }>(`/api/hods/combinations?${q.toString()}`)
      .then((r) => setCombos(r.combinations || []))
      .catch(() => setCombos([]));
  };

  useEffect(() => {
    loadCombos();
  }, [comboLocationSearch, comboDepartmentSearch]);

  const save = async () => {
    if (editingId) {
      await api(`/api/hods/${editingId}`, {
        method: "PUT",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ location, department, hodEmail, hodName: "" }),
      });
    } else {
      await api("/api/hods", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ location, department, hodEmail, hodName: "" }),
      });
    }
    setLocation("");
    setDepartment("");
    setHodEmail("");
    setEditingId(null);
    load();
  };

  const startEdit = (item: any) => {
    setEditingId(item.Id);
    setLocation(item.Location || "");
    setDepartment(item.Department || "");
    setHodEmail(item.HodEmail || "");
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const remove = async (id: number) => {
    if (!window.confirm("Deactivate this HOD mapping?")) return;
    await api(`/api/hods/${id}`, { method: "DELETE" });
    if (editingId === id) {
      setEditingId(null);
      setLocation("");
      setDepartment("");
      setHodEmail("");
    }
    load();
  };

  return (
    <div className="grid xl:grid-cols-3 gap-4 animate-fade-in">
      <Card className="glass xl:col-span-2" data-tour="hod-combinations">
        <CardHeader>
          <CardTitle>HOD Matrix Master</CardTitle>
          <CardDescription>Pick exact Location + Department combinations first, then assign HOD.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-2">
          <div className="grid md:grid-cols-2 gap-2">
            <Input placeholder="Search combinations by location" value={comboLocationSearch} onChange={(e) => setComboLocationSearch(e.target.value)} />
            <Input placeholder="Search combinations by department" value={comboDepartmentSearch} onChange={(e) => setComboDepartmentSearch(e.target.value)} />
          </div>

          <div className="max-h-52 overflow-auto border rounded-md text-sm">
            {combos.map((c, idx) => (
              <button
                key={`${c.Location}-${c.Department}-${idx}`}
                type="button"
                className="w-full text-left p-2 border-b hover:bg-muted"
                onClick={() => {
                  setLocation(c.Location);
                  setDepartment(c.Department);
                }}
              >
                <strong>{c.Location}</strong> / {c.Department}
              </button>
            ))}
          </div>

          <div className="grid md:grid-cols-2 gap-2">
            <Input placeholder="Location" value={location} onChange={(e) => setLocation(e.target.value)} />
            <Input placeholder="Department" value={department} onChange={(e) => setDepartment(e.target.value)} />
          </div>
          <Input placeholder="Search employee" value={search} onChange={(e) => setSearch(e.target.value)} />
          <div className="max-h-28 overflow-auto border rounded-md text-sm">
            {employees.map((e) => (
              <button
                type="button"
                key={e.EmpEmail}
                className="w-full text-left p-2 border-b hover:bg-muted"
                onClick={() => setHodEmail(e.EmpEmail)}
              >
                {e.EmpName} ({e.EmpEmail})
              </button>
            ))}
          </div>
          <Input placeholder="Selected HOD Email" value={hodEmail} onChange={(e) => setHodEmail(e.target.value)} />
          <div className="flex gap-2">
            <Button onClick={save}>{editingId ? "Update Mapping" : "Save Mapping"}</Button>
            {editingId ? (
              <Button
                variant="outline"
                onClick={() => {
                  setEditingId(null);
                  setLocation("");
                  setDepartment("");
                  setHodEmail("");
                }}
              >
                Cancel Edit
              </Button>
            ) : null}
          </div>
        </CardContent>
      </Card>

      <Card className="glass" data-tour="analytics-main">
        <CardHeader><CardTitle>Configured HODs</CardTitle></CardHeader>
        <CardContent className="max-h-[560px] overflow-auto">
          {list.map((r) => (
            <div key={r.Id} className="p-2 border-b text-sm flex items-center justify-between">
              <div>
                <div className="font-medium">{r.Location} / {r.Department}</div>
                <div className="text-muted-foreground">{r.HodEmail}</div>
              </div>
              <div className="flex items-center gap-2">
                <Badge variant={r.Active ? "success" : "secondary"}>{r.Active ? "Active" : "Inactive"}</Badge>
                <Button size="sm" variant="outline" onClick={() => startEdit(r)}>Edit</Button>
                <Button size="sm" variant="destructive" onClick={() => remove(r.Id)}>Delete</Button>
              </div>
            </div>
          ))}
        </CardContent>
      </Card>
    </div>
  );
}

function AnalyticsPage() {
  const [overview, setOverview] = useState<any>(null);
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
        setOverview(o);
        setByDept(d);
        setByLoc(l);
        setByUser(u);
      })
      .catch(() => undefined);
  }, []);

  const stats = useMemo(() => {
    if (!overview) return [];
    return [
      { name: "Total", value: overview.totalDocuments || 0 },
      { name: "Controlled", value: overview.controlledDocuments || 0 },
      { name: "Approved", value: overview.approvedDocuments || 0 },
      { name: "Pending", value: overview.pendingApprovals || 0 },
    ];
  }, [overview]);

  return (
    <div className="space-y-4 animate-fade-in">
      <div className="grid md:grid-cols-4 gap-3">
        <Stat title="Total Documents" value={String(overview?.totalDocuments || 0)} icon={<Activity className="h-5 w-5" />} />
        <Stat title="Controlled" value={String(overview?.controlledDocuments || 0)} icon={<ShieldCheck className="h-5 w-5" />} />
        <Stat title="Pending Approvals" value={String(overview?.pendingApprovals || 0)} icon={<Clock3 className="h-5 w-5" />} />
        <Stat title="Uploaders" value={String(overview?.uniqueUploaders || 0)} icon={<Users className="h-5 w-5" />} />
      </div>

      <Card className="glass">
        <CardHeader><CardTitle>Document Distribution</CardTitle></CardHeader>
        <CardContent className="h-[320px]">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={byDept}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="Department" />
              <YAxis />
              <Tooltip />
              <Bar dataKey="DocumentCount" fill="#0284c7" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </CardContent>
      </Card>

      <div className="grid xl:grid-cols-2 gap-4">
        <Card className="glass">
          <CardHeader><CardTitle>By Location</CardTitle></CardHeader>
          <CardContent className="h-[300px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={byLoc}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="Location" />
                <YAxis />
                <Tooltip />
                <Bar dataKey="DocumentCount" fill="#16a34a" radius={[8, 8, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>

        <Card className="glass">
          <CardHeader><CardTitle>Controlled Mix</CardTitle></CardHeader>
          <CardContent className="h-[300px]">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie data={stats} dataKey="value" nameKey="name" outerRadius={100} label>
                  {stats.map((_, i) => (
                    <Cell key={i} fill={["#0284c7", "#f59e0b", "#22c55e", "#ef4444"][i % 4]} />
                  ))}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
      </div>

      <Card className="glass">
        <CardHeader><CardTitle>Top Uploaders</CardTitle></CardHeader>
        <CardContent>
          <div className="max-h-48 overflow-auto">
            {byUser.map((u) => (
              <div className="p-2 border-b text-sm" key={u.CreatorEmail}>{u.CreatorEmail} · {u.DocumentCount}</div>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

function AuditPage() {
  const [logs, setLogs] = useState<any[]>([]);
  const [action, setAction] = useState("");
  const [user, setUser] = useState("");

  const load = () => {
    const q = new URLSearchParams();
    if (action) q.set("action", action);
    if (user) q.set("user", user);
    q.set("pageSize", "100");
    api<{ logs: any[] }>(`/api/audit-log?${q.toString()}`).then((r) => setLogs(r.logs || [])).catch(() => setLogs([]));
  };

  useEffect(() => {
    load();
  }, []);

  return (
    <Card className="glass animate-fade-in">
      <CardHeader>
        <CardTitle>Audit Log</CardTitle>
        <CardDescription>Timestamp, actor, action, reason, before state, after state and commit trail.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-3">
        <div className="grid md:grid-cols-3 gap-2" data-tour="audit-filters">
          <Input placeholder="Action" value={action} onChange={(e) => setAction(e.target.value)} />
          <Input placeholder="User email" value={user} onChange={(e) => setUser(e.target.value)} />
          <Button onClick={load}>Filter</Button>
        </div>
        <div className="max-h-[540px] overflow-auto border rounded-md">
          {logs.map((log) => (
            <div className="p-3 border-b text-sm space-y-1" key={log.Id}>
              <div className="font-semibold">{log.Action}</div>
              <div className="text-muted-foreground">{log.UserEmail || "System"} · {new Date(log.CreatedAt).toLocaleString()}</div>
              {log.Reason ? <div><strong>Reason:</strong> {log.Reason}</div> : null}
              <div className="grid md:grid-cols-2 gap-2 text-xs">
                <pre className="p-2 bg-muted rounded overflow-auto">Before: {log.BeforeState || "-"}</pre>
                <pre className="p-2 bg-muted rounded overflow-auto">After: {log.AfterState || "-"}</pre>
              </div>
            </div>
          ))}
        </div>
      </CardContent>
    </Card>
  );
}

function PublicViewerPage() {
  const { token = "" } = useParams();
  const [meta, setMeta] = useState<any>(null);
  const [error, setError] = useState("");

  useEffect(() => {
    api(`/api/public/meta/${token}`)
      .then((data) => {
        setMeta(data);
        setError("");
      })
      .catch((e) => setError(String(e.message || "Invalid or expired link")));
  }, [token]);

  return (
    <div className="min-h-screen bg-page-gradient no-print-view p-4 md:p-8">
      <div className="max-w-6xl mx-auto space-y-3">
        <Card className="glass">
          <CardHeader>
            <CardTitle className="font-display text-4xl tracking-wide">Public View</CardTitle>
            <CardDescription>View-only mode. Download and print are disabled by policy.</CardDescription>
          </CardHeader>
          <CardContent>
            {error ? <p className="text-destructive">{error}</p> : null}
            {meta ? (
              <div className="space-y-2 text-sm">
                <p><strong>Title:</strong> {meta.Title}</p>
                <p><strong>Version:</strong> {meta.CurrentVersion}</p>
                <p><strong>Controlled:</strong> {String(Boolean(meta.IsControlled))}</p>
              </div>
            ) : null}
          </CardContent>
        </Card>
        <iframe title="public-doc" className="w-full h-[78vh] rounded-xl border bg-white" src={`/api/public/view/${token}`} />
      </div>
    </div>
  );
}

function Stat({ title, value, icon }: { title: string; value: string; icon: React.ReactNode }) {
  return (
    <Card className="glass hover:scale-[1.02] transition-transform duration-300">
      <CardContent className="pt-6 flex items-center justify-between">
        <div>
          <p className="text-xs text-muted-foreground">{title}</p>
          <p className="text-2xl font-bold">{value}</p>
        </div>
        <div className="h-10 w-10 rounded-xl bg-gradient-to-br from-sky-500/20 to-emerald-500/20 flex items-center justify-center">
          {icon}
        </div>
      </CardContent>
    </Card>
  );
}

function NotAdmin() {
  return (
    <Card>
      <CardContent className="pt-6">This section is restricted to admin users.</CardContent>
    </Card>
  );
}

function AdminGate({ children }: { children: React.ReactNode }) {
  const { user } = useAuth();
  if (!user?.isAdmin) return <NotAdmin />;
  return <>{children}</>;
}

function AuthenticatedApp() {
  const { loading, error, user } = useAuth();
  if (loading) return <LoadingScreen />;
  if (error) return <LoadingScreen />;
  if (!user) return <LoadingScreen />;

  return (
    <ProtectedLayout>
      <Routes>
        <Route path="/" element={<DashboardPage />} />
        <Route path="/upload" element={<UploadPage />} />
        <Route path="/approvals" element={<ApprovalsPage />} />
        <Route path="/verify" element={<VerifyPage />} />
        <Route path="/admin/users" element={<AdminGate><AdminUsersPage /></AdminGate>} />
        <Route path="/admin/hods" element={<AdminGate><HodsPage /></AdminGate>} />
        <Route path="/admin/analytics" element={<AdminGate><AnalyticsPage /></AdminGate>} />
        <Route path="/admin/audit" element={<AdminGate><AuditPage /></AdminGate>} />
        <Route path="*" element={<Card><CardContent className="pt-6"><Sparkles className="h-4 w-4 inline mr-2" />Route not found. Go to <Link className="text-primary underline" to="/">Dashboard</Link>.</CardContent></Card>} />
      </Routes>
    </ProtectedLayout>
  );
}

export default function App() {
  return (
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
  );
}
