import React, { useState, useEffect, useMemo } from "react";
import {
  LayoutDashboard,
  Database,
  PlusCircle,
  FileText,
  LogOut,
  ShieldAlert,
  Activity,
  CheckCircle2,
  Clock,
  Trash2,
  Edit,
  Download,
  Upload,
  Search,
  BrainCircuit,
  Loader2,
  X,
  Calendar,
  ChevronDown,
} from "lucide-react";
import * as XLSX from "xlsx";
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
  Legend,
} from "recharts";
import { dbService } from "./db";
import { analyzeCase } from "./services/geminiService";
import {
  Page,
  CaseRecord,
  CaseStatus,
  Student,
  Teacher,
  DatabaseState,
  Attachment,
} from "./types";

// Components
const Navbar: React.FC<{
  currentPage: Page;
  onPageChange: (p: Page) => void;
  onLogout: () => void;
}> = ({ currentPage, onPageChange, onLogout }) => (
  <nav className="sticky top-0 z-50 bg-white/80 backdrop-blur-xl border-b border-slate-200 px-6 py-4 flex justify-between items-center shadow-sm">
    <div className="flex items-center gap-4">
      <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white font-bold text-xl shadow-md shadow-indigo-200">
        BK
      </div>
      <h1 className="font-bold text-lg text-slate-800 leading-none uppercase hidden md:block">
        BK Peduli Siswa
      </h1>
    </div>
    <div className="flex gap-2 md:gap-4 text-xs md:text-sm font-semibold overflow-x-auto">
      <button
        onClick={() => onPageChange("dashboard")}
        className={`px-3 py-2 rounded-xl transition flex items-center gap-2 ${currentPage === "dashboard" ? "bg-slate-100 text-indigo-600" : "hover:bg-slate-50 text-slate-600"}`}
      >
        <LayoutDashboard size={18} />{" "}
        <span className="hidden sm:inline">Beranda</span>
      </button>
      <button
        onClick={() => onPageChange("master")}
        className={`px-3 py-2 rounded-xl transition flex items-center gap-2 ${currentPage === "master" ? "bg-slate-100 text-indigo-600" : "hover:bg-slate-50 text-slate-600"}`}
      >
        <Database size={18} /> <span className="hidden sm:inline">Master</span>
      </button>
      <button
        onClick={() => onPageChange("input")}
        className={`px-4 py-2 rounded-xl transition flex items-center gap-2 whitespace-nowrap ${currentPage === "input" ? "bg-indigo-700 text-white shadow-md shadow-indigo-200" : "bg-indigo-600 text-white shadow-md shadow-indigo-100 hover:bg-indigo-700"}`}
      >
        <PlusCircle size={18} /> <span>Input Kasus</span>
      </button>
      <button
        onClick={() => onPageChange("laporan")}
        className={`px-3 py-2 rounded-xl transition flex items-center gap-2 ${currentPage === "laporan" ? "bg-slate-100 text-indigo-600" : "hover:bg-slate-50 text-slate-600"}`}
      >
        <FileText size={18} /> <span className="hidden sm:inline">Laporan</span>
      </button>
      <button
        onClick={onLogout}
        className="px-3 py-2 rounded-xl text-red-500 hover:bg-red-50 transition flex items-center gap-2"
      >
        <LogOut size={18} /> <span className="hidden sm:inline">Keluar</span>
      </button>
    </div>
  </nav>
);

const App: React.FC = () => {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [currentPage, setCurrentPage] = useState<Page>("dashboard");
  const [dbState, setDbState] = useState<DatabaseState>({
    siswa: [],
    wali_kelas: [],
    guru_bk: [],
    kasus: [],
  });
  const [editingCase, setEditingCase] = useState<CaseRecord | null>(null);
  const [loginCreds, setLoginCreds] = useState({ user: "", pass: "" });
  const [loginError, setLoginError] = useState(false);
  const [aiAnalysis, setAiAnalysis] = useState<any>(null);
  const [isAiLoading, setIsAiLoading] = useState(false);

  // Initialize data
  useEffect(() => {
    const loadData = async () => {
      await dbService.init();
      const siswa = await dbService.getAll<Student>("siswa");
      const wali_kelas = await dbService.getAll<Teacher>("wali_kelas");
      const guru_bk = await dbService.getAll<Teacher>("guru_bk");
      const kasus = await dbService.getAll<CaseRecord>("kasus");
      setDbState({ siswa, wali_kelas, guru_bk, kasus });
    };
    loadData();
  }, []);

  const refreshMemory = async () => {
    const siswa = await dbService.getAll<Student>("siswa");
    const wali_kelas = await dbService.getAll<Teacher>("wali_kelas");
    const guru_bk = await dbService.getAll<Teacher>("guru_bk");
    const kasus = await dbService.getAll<CaseRecord>("kasus");
    setDbState({ siswa, wali_kelas, guru_bk, kasus });
  };

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (loginCreds.user === "admin" && loginCreds.pass === "admin123") {
      setIsLoggedIn(true);
      setLoginError(false);
    } else {
      setLoginError(true);
    }
  };

  const handleLogout = () => {
    if (confirm("Yakin ingin keluar?")) {
      setIsLoggedIn(false);
      setCurrentPage("dashboard");
    }
  };

  const handleCaseSubmit = async (record: Omit<CaseRecord, "created_at">) => {
    const newRecord = {
      ...record,
      created_at: editingCase ? editingCase.created_at : Date.now(),
    };
    await dbService.put("kasus", newRecord);
    await refreshMemory();
    setEditingCase(null);
    setCurrentPage("laporan");
    alert("Data berhasil disimpan!");
  };

  const handleDeleteCase = async (id: number) => {
    if (confirm("Apakah Anda yakin ingin menghapus data ini?")) {
      await dbService.delete("kasus", id);
      await refreshMemory();
    }
  };

  const handleEditCase = (record: CaseRecord) => {
    setEditingCase(record);
    setCurrentPage("input");
  };

  const handleStatusUpdate = async (id: number, status: CaseStatus) => {
    const item = dbState.kasus.find((k) => k.id === id);
    if (item) {
      await dbService.put("kasus", { ...item, status });
      await refreshMemory();
    }
  };

  const handleAiAnalysis = async (kronologi: string, kategori: string) => {
    setIsAiLoading(true);
    setAiAnalysis(null);
    try {
      const result = await analyzeCase(kronologi, kategori);
      setAiAnalysis(result);
    } catch (error: any) {
      if (error.message.includes("GEMINI_API_KEY")) {
        alert("KONFIGURASI ERROR: GEMINI_API_KEY tidak ditemukan. Silakan atur di Environment Variables pada pengaturan proyek Vercel Anda.");
      } else {
        alert("Terjadi kesalahan saat menghubungi layanan AI. Coba lagi nanti.");
        console.error("AI Analysis error:", error);
      }
    } finally {
      setIsAiLoading(false);
    }
  };

  // Views
  if (!isLoggedIn) {
    return (
      <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-50">
        <div className="bg-white w-full max-w-md rounded-3xl shadow-xl p-10 border border-slate-100">
          <div className="text-center mb-8">
            <div className="w-16 h-16 bg-indigo-600 rounded-2xl mx-auto mb-6 flex items-center justify-center text-white text-2xl font-bold shadow-lg shadow-indigo-200">
              BK
            </div>
            <h2 className="text-2xl font-bold text-slate-900 tracking-tight">
              BK Peduli Siswa
            </h2>
            <p className="text-slate-500 text-sm mt-2">
              Pendekatan Edukatif dan Humanis
            </p>
          </div>
          <form onSubmit={handleLogin} className="space-y-5">
            <div className="space-y-1.5">
              <label className="block text-xs font-semibold uppercase text-slate-500 ml-1">
                Username
              </label>
              <input
                type="text"
                required
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-slate-50 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                placeholder="admin"
                value={loginCreds.user}
                onChange={(e) =>
                  setLoginCreds({ ...loginCreds, user: e.target.value })
                }
              />
            </div>
            <div className="space-y-1.5">
              <label className="block text-xs font-semibold uppercase text-slate-500 ml-1">
                Password
              </label>
              <input
                type="password"
                required
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-slate-50 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                placeholder="••••••••"
                value={loginCreds.pass}
                onChange={(e) =>
                  setLoginCreds({ ...loginCreds, pass: e.target.value })
                }
              />
            </div>
            {loginError && (
              <div className="text-red-500 text-xs font-medium text-center bg-red-50 py-2 rounded-lg">
                Username atau Password salah!
              </div>
            )}
            <button
              type="submit"
              className="w-full bg-indigo-600 text-white py-3.5 rounded-xl font-semibold hover:bg-indigo-700 transition duration-200 mt-2 shadow-md shadow-indigo-200 flex items-center justify-center gap-2"
            >
              <ShieldAlert size={18} /> Buka Layanan
            </button>
          </form>
          <p className="text-center text-[10px] text-slate-400 mt-8 uppercase tracking-widest font-medium">
            Database Local Storage Terenkripsi
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen pb-20">
      <Navbar
        currentPage={currentPage}
        onPageChange={(p) => {
          if (p === "input" && !editingCase) {
            setEditingCase(null);
          }
          setCurrentPage(p);
        }}
        onLogout={handleLogout}
      />

      <main className="max-w-7xl mx-auto p-4 md:p-8">
        {currentPage === "dashboard" && <DashboardView dbState={dbState} />}
        {currentPage === "master" && (
          <MasterView dbState={dbState} onRefresh={refreshMemory} />
        )}
        {currentPage === "input" && (
          <InputView
            dbState={dbState}
            editingCase={editingCase}
            onSubmit={handleCaseSubmit}
            onCancel={() => {
              setEditingCase(null);
              setCurrentPage("laporan");
            }}
            onAiAnalysis={handleAiAnalysis}
            onClearAiAnalysis={() => setAiAnalysis(null)}
            aiAnalysis={aiAnalysis}
            isAiLoading={isAiLoading}
          />
        )}
        {currentPage === "laporan" && (
          <ReportView
            dbState={dbState}
            onEdit={handleEditCase}
            onDelete={handleDeleteCase}
            onStatusUpdate={handleStatusUpdate}
          />
        )}
      </main>
    </div>
  );
};

// Sub-Views
const DashboardView: React.FC<{ dbState: DatabaseState }> = ({ dbState }) => {
  const stats = useMemo(
    () => ({
      baru: dbState.kasus.filter((k) => k.status === CaseStatus.BARU).length,
      proses: dbState.kasus.filter((k) => k.status === CaseStatus.PROSES)
        .length,
      selesai: dbState.kasus.filter((k) => k.status === CaseStatus.SELESAI)
        .length,
    }),
    [dbState.kasus],
  );

  const chartData = [
    { name: "Baru", count: stats.baru, color: "#3b82f6" },
    { name: "Proses", count: stats.proses, color: "#f59e0b" },
    { name: "Selesai", count: stats.selesai, color: "#10b981" },
  ];

  const categoryData = useMemo(() => {
    const counts: Record<string, number> = {};
    dbState.kasus.forEach((k) => {
      counts[k.kategori_kasus] = (counts[k.kategori_kasus] || 0) + 1;
    });
    return Object.entries(counts).map(([name, value]) => ({ name, value }));
  }, [dbState.kasus]);

  const COLORS = [
    "#6366f1",
    "#a855f7",
    "#ec4899",
    "#f43f5e",
    "#10b981",
    "#f59e0b",
  ];

  return (
    <div className="space-y-6 md:space-y-8 animate-in fade-in duration-500">
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 md:gap-6">
        <div className="glass-card p-5 md:p-6 flex justify-between items-center group hover:shadow-md transition-all">
          <div>
            <p className="text-[10px] md:text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Antrean Kasus
            </p>
            <h2 className="text-3xl md:text-4xl font-light mt-1 md:mt-2 text-slate-800">
              {stats.baru}
            </h2>
          </div>
          <div className="w-10 h-10 md:w-12 md:h-12 rounded-full bg-blue-50 flex items-center justify-center text-blue-500">
            <Clock size={20} className="md:w-6 md:h-6" />
          </div>
        </div>
        <div className="glass-card p-5 md:p-6 flex justify-between items-center group hover:shadow-md transition-all">
          <div>
            <p className="text-[10px] md:text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Dalam Proses
            </p>
            <h2 className="text-3xl md:text-4xl font-light mt-1 md:mt-2 text-slate-800">
              {stats.proses}
            </h2>
          </div>
          <div className="w-10 h-10 md:w-12 md:h-12 rounded-full bg-amber-50 flex items-center justify-center text-amber-500">
            <Activity size={20} className="md:w-6 md:h-6" />
          </div>
        </div>
        <div className="glass-card p-5 md:p-6 flex justify-between items-center group hover:shadow-md transition-all">
          <div>
            <p className="text-[10px] md:text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Selesai
            </p>
            <h2 className="text-3xl md:text-4xl font-light mt-1 md:mt-2 text-slate-800">
              {stats.selesai}
            </h2>
          </div>
          <div className="w-10 h-10 md:w-12 md:h-12 rounded-full bg-emerald-50 flex items-center justify-center text-emerald-500">
            <CheckCircle2 size={20} className="md:w-6 md:h-6" />
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 md:gap-8">
        <div className="glass-card p-5 md:p-8 min-h-[350px] md:min-h-[400px]">
          <h3 className="text-sm md:text-base font-semibold mb-6 md:mb-8 text-slate-800 flex items-center gap-2">
            <LayoutDashboard size={18} className="text-indigo-500" /> Statistik
            Penanganan
          </h3>
          <div className="h-[250px] md:h-[300px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData}>
                <CartesianGrid
                  strokeDasharray="3 3"
                  vertical={false}
                  stroke="#f1f5f9"
                />
                <XAxis
                  dataKey="name"
                  axisLine={false}
                  tickLine={false}
                  tick={{ fill: "#64748b", fontSize: 12 }}
                />
                <YAxis
                  axisLine={false}
                  tickLine={false}
                  tick={{ fill: "#64748b", fontSize: 12 }}
                />
                <Tooltip
                  cursor={{ fill: "#f8fafc" }}
                  contentStyle={{
                    borderRadius: "12px",
                    border: "none",
                    boxShadow: "0 10px 15px -3px rgba(0,0,0,0.1)",
                  }}
                />
                <Bar dataKey="count" radius={[8, 8, 0, 0]}>
                  {chartData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="glass-card p-5 md:p-8 min-h-[350px] md:min-h-[400px]">
          <h3 className="text-sm md:text-base font-semibold mb-6 md:mb-8 text-slate-800 flex items-center gap-2">
            <ShieldAlert size={18} className="text-indigo-500" /> Distribusi
            Kategori Kasus
          </h3>
          <div className="h-[250px] md:h-[300px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={categoryData}
                  innerRadius={60}
                  outerRadius={80}
                  paddingAngle={5}
                  dataKey="value"
                >
                  {categoryData.map((entry, index) => (
                    <Cell
                      key={`cell-${index}`}
                      fill={COLORS[index % COLORS.length]}
                    />
                  ))}
                </Pie>
                <Tooltip
                  contentStyle={{
                    borderRadius: "12px",
                    border: "none",
                    boxShadow: "0 10px 15px -3px rgba(0,0,0,0.1)",
                  }}
                />
                <Legend
                  layout="vertical"
                  align="right"
                  verticalAlign="middle"
                  iconType="circle"
                />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    </div>
  );
};

const MasterView: React.FC<{
  dbState: DatabaseState;
  onRefresh: () => void;
}> = ({ dbState, onRefresh }) => {
  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: "binary" });

      const sheets = {
        Siswa: "siswa",
        WaliKelas: "wali_kelas",
        GuruBK: "guru_bk",
      };

      for (const [sheetName, storeName] of Object.entries(sheets)) {
        const ws = wb.Sheets[sheetName];
        if (ws) {
          const data = XLSX.utils.sheet_to_json(ws);
          await dbService.clear(storeName);
          for (const item of data as any[]) {
            await dbService.put(storeName, item);
          }
        }
      }

      onRefresh();
      alert("Data Master berhasil diimpor!");
    };
    reader.readAsBinaryString(file);
  };

  const handleBackup = () => {
    const data = JSON.stringify(dbState);
    const blob = new window.Blob([data], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Backup_BK_${new Date().toISOString().split("T")[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleRestore = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data: DatabaseState = JSON.parse(evt.target?.result as string);
        const stores = ["siswa", "wali_kelas", "guru_bk", "kasus"];
        for (const store of stores) {
          await dbService.clear(store);
          const items = (data as any)[store] || [];
          for (const item of items as any[]) {
            await dbService.put(store, item);
          }
        }
        onRefresh();
        alert("Database berhasil direstore!");
      } catch (err) {
        alert("File backup tidak valid!");
      }
    };
    reader.readAsText(file);
  };

  return (
    <div className="space-y-6 animate-in slide-in-from-bottom-4 duration-500">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 md:gap-6">
        <div className="glass-card p-5 md:p-8">
          <h3 className="text-base md:text-lg font-semibold text-slate-800 mb-2 flex items-center gap-2">
            <Upload size={20} className="text-indigo-500" /> Import Data Master
            (Excel)
          </h3>
          <p className="text-[10px] md:text-xs text-slate-500 mb-4 bg-slate-50 p-3 md:p-4 rounded-xl border border-slate-100">
            Gunakan file Excel dengan Sheet: <b>Siswa</b>, <b>WaliKelas</b>,{" "}
            <b>GuruBK</b>.<br />
            Kolom Siswa: Nama, Kelas. Kolom lainnya: Nama.
            <span className="text-red-500 font-medium mt-1 md:mt-2 block">
              *Data lama akan ditimpa!
            </span>
          </p>
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleExcelImport}
            className="block w-full text-xs md:text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 md:file:py-2.5 md:file:px-6 file:rounded-xl file:border-0 file:text-[10px] md:file:text-xs file:font-semibold file:bg-indigo-50 file:text-indigo-600 cursor-pointer hover:file:bg-indigo-100 transition-all"
          />
        </div>
        <div className="glass-card p-5 md:p-8">
          <h3 className="text-base md:text-lg font-semibold text-slate-800 mb-2 flex items-center gap-2">
            <Database size={20} className="text-indigo-500" /> Backup & Restore
          </h3>
          <p className="text-[10px] md:text-xs text-slate-500 mb-4">
            Amankan seluruh data kasus dan data master ke file JSON.
          </p>
          <div className="flex flex-col sm:flex-row gap-3 md:gap-4 mt-2">
            <button
              onClick={handleBackup}
              className="flex-1 bg-slate-800 text-white py-2.5 md:py-3 rounded-xl text-[10px] md:text-xs font-semibold hover:bg-slate-700 transition flex items-center justify-center gap-2"
            >
              <Download size={16} /> Backup DB
            </button>
            <label className="flex-1 bg-white border border-slate-200 text-slate-700 py-2.5 md:py-3 rounded-xl text-[10px] md:text-xs font-semibold text-center cursor-pointer hover:bg-slate-50 transition flex items-center justify-center gap-2">
              <Upload size={16} /> Restore{" "}
              <input
                type="file"
                accept=".json"
                className="hidden"
                onChange={handleRestore}
              />
            </label>
          </div>
        </div>
      </div>

      <div className="glass-card p-5 md:p-8">
        <h4 className="font-semibold text-xs md:text-sm mb-4 md:mb-6 text-slate-500">
          Preview Ketersediaan Data
        </h4>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 md:gap-4 text-center">
          <div className="bg-slate-50 p-4 md:p-6 rounded-2xl border border-slate-100 group hover:bg-slate-100 transition-colors">
            <span className="block text-2xl md:text-3xl font-light text-slate-800 mb-1">
              {dbState.siswa.length}
            </span>
            <span className="text-[10px] md:text-xs font-medium text-slate-500">
              Siswa Terdaftar
            </span>
          </div>
          <div className="bg-slate-50 p-4 md:p-6 rounded-2xl border border-slate-100 group hover:bg-slate-100 transition-colors">
            <span className="block text-2xl md:text-3xl font-light text-slate-800 mb-1">
              {dbState.wali_kelas.length}
            </span>
            <span className="text-[10px] md:text-xs font-medium text-slate-500">
              Wali Kelas
            </span>
          </div>
          <div className="bg-slate-50 p-4 md:p-6 rounded-2xl border border-slate-100 group hover:bg-slate-100 transition-colors">
            <span className="block text-2xl md:text-3xl font-light text-slate-800 mb-1">
              {dbState.guru_bk.length}
            </span>
            <span className="text-[10px] md:text-xs font-medium text-slate-500">Guru BK</span>
          </div>
        </div>
      </div>
    </div>
  );
};

const InputView: React.FC<{
  dbState: DatabaseState;
  editingCase: CaseRecord | null;
  onSubmit: (r: any) => void;
  onCancel: () => void;
  onAiAnalysis: (kron: string, kat: string) => void;
  onClearAiAnalysis: () => void;
  aiAnalysis: any;
  isAiLoading: boolean;
}> = ({
  dbState,
  editingCase,
  onSubmit,
  onCancel,
  onAiAnalysis,
  onClearAiAnalysis,
  aiAnalysis,
  isAiLoading,
}) => {
  const [formData, setFormData] = useState<Partial<CaseRecord>>({
    tanggal: new Date().toISOString().split("T")[0],
    kategori_kasus: "",
    kelas: "",
    nama_siswa: "",
    guru_kelas: "",
    guru_bk: "",
    kronologi: "",
    tindak_lanjut: "Konseling Individu",
    status: CaseStatus.BARU,
    lampiran: [],
  });

  useEffect(() => {
    if (editingCase) setFormData(editingCase);
  }, [editingCase]);

  const CLASSES = useMemo(() => {
    const list: string[] = [];
    ["7", "8", "9"].forEach((grade) => {
      ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((section) => {
        list.push(grade + section);
      });
    });
    return list;
  }, []);

  // Updated to handle case-insensitive Excel properties (Nama/nama, Kelas/kelas)
  const filteredStudents = useMemo(() => {
    if (!formData.kelas) return [];
    return dbState.siswa.filter((s) => {
      const sKelas = (s as any).Kelas || (s as any).kelas;
      return (
        sKelas &&
        sKelas.toString().toUpperCase() === formData.kelas?.toUpperCase()
      );
    });
  }, [formData.kelas, dbState.siswa]);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    files.forEach((f: any) => {
      const reader = new FileReader();
      reader.onload = (ev) => {
        setFormData((prev) => ({
          ...prev,
          lampiran: [
            ...(prev.lampiran || []),
            {
              name: f.name,
              type: f.type,
              data: ev.target?.result as string,
            },
          ],
        }));
      };
      reader.readAsDataURL(f);
    });
  };

  const removeFile = (idx: number) => {
    setFormData((prev) => ({
      ...prev,
      lampiran: prev.lampiran?.filter((_, i) => i !== idx),
    }));
  };

  const inputClass =
    "w-full p-3 md:p-4 rounded-xl border border-slate-200 bg-slate-50 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-400 text-slate-800 appearance-none text-sm md:text-base";

  return (
    <div className="max-w-5xl mx-auto animate-in zoom-in-95 duration-300">
      <form
        onSubmit={(e) => {
          e.preventDefault();
          onSubmit(formData);
        }}
        className="glass-card p-5 md:p-8 lg:p-12 relative overflow-hidden"
      >
        <div className="absolute top-0 left-0 w-full h-1.5 bg-indigo-500"></div>

        {editingCase && (
          <div className="absolute top-5 right-5 md:top-8 md:right-12 bg-amber-50 text-amber-700 px-3 py-1 md:px-4 md:py-1.5 rounded-full text-[10px] md:text-xs font-semibold uppercase tracking-wider border border-amber-200 animate-pulse">
            Edit Mode
          </div>
        )}

        <div className="mb-6 md:mb-10">
          <h3 className="text-2xl md:text-3xl font-bold text-slate-900 tracking-tight">
            Input Detail Kasus
          </h3>
          <p className="text-slate-500 text-sm md:text-base mt-1 md:mt-2">
            Lengkapi data laporan siswa secara objektif dan akurat.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 md:gap-x-12 gap-y-6 md:gap-y-12">
          {/* Row 1 */}
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              TANGGAL KEJADIAN
            </label>
            <div className="relative group">
              <input
                type="date"
                required
                className={inputClass + " pr-10 md:pr-12"}
                value={formData.tanggal}
                onChange={(e) =>
                  setFormData({ ...formData, tanggal: e.target.value })
                }
              />
              <Calendar
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-indigo-500 transition-colors pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
          </div>

          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              KATEGORI KASUS
            </label>
            <div className="relative">
              <select
                required
                className={inputClass}
                value={formData.kategori_kasus}
                onChange={(e) =>
                  setFormData({ ...formData, kategori_kasus: e.target.value })
                }
              >
                <option value="">-- Pilih Jenis Kasus --</option>
                {[
                  "Kedisiplinan",
                  "Etika",
                  "Akademik",
                  "Bullying",
                  "Perkelahian",
                  "Merokok",
                  "Narkoba",
                  "Lainnya",
                ].map((k) => (
                  <option key={k} value={k}>
                    {k}
                  </option>
                ))}
              </select>
              <ChevronDown
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
          </div>

          {/* Row 2 */}
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              KELAS
            </label>
            <div className="relative">
              <select
                required
                className={inputClass}
                value={formData.kelas}
                onChange={(e) =>
                  setFormData({
                    ...formData,
                    kelas: e.target.value,
                    nama_siswa: "",
                  })
                }
              >
                <option value="">-- Pilih Kelas --</option>
                {CLASSES.map((k) => (
                  <option key={k} value={k}>
                    {k}
                  </option>
                ))}
              </select>
              <ChevronDown
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
          </div>

          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              NAMA SISWA
            </label>
            <div className="relative">
              {/* Changed from Input list to Select for better visibility of data */}
              <select
                required
                className={inputClass}
                value={formData.nama_siswa}
                onChange={(e) =>
                  setFormData({ ...formData, nama_siswa: e.target.value })
                }
                disabled={!formData.kelas}
              >
                {!formData.kelas ? (
                  <option value="">Pilih kelas terlebih dahulu...</option>
                ) : filteredStudents.length === 0 ? (
                  <option value="">
                    Tidak ada data siswa di kelas ini (Cek Master)
                  </option>
                ) : (
                  <>
                    <option value="">-- Pilih Nama Siswa --</option>
                    {filteredStudents.map((s, i) => {
                      const sNama = (s as any).Nama || (s as any).nama;
                      return (
                        <option key={i} value={sNama}>
                          {sNama}
                        </option>
                      );
                    })}
                  </>
                )}
              </select>
              <ChevronDown
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
          </div>

          {/* Row 3 */}
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              WALI KELAS
            </label>
            <div className="relative">
              <input
                list="wali-list"
                className={inputClass}
                placeholder="Masukkan nama Wali Kelas"
                value={formData.guru_kelas}
                onChange={(e) =>
                  setFormData({ ...formData, guru_kelas: e.target.value })
                }
              />
              <datalist id="wali-list">
                {dbState.wali_kelas.map((t, i) => (
                  <option key={i} value={t.Nama || (t as any).nama} />
                ))}
              </datalist>
            </div>
          </div>

          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              GURU BK
            </label>
            <div className="relative">
              <input
                list="bk-list"
                required
                className={inputClass}
                placeholder="Masukkan nama Guru BK"
                value={formData.guru_bk}
                onChange={(e) =>
                  setFormData({ ...formData, guru_bk: e.target.value })
                }
              />
              <datalist id="bk-list">
                {dbState.guru_bk.map((t, i) => (
                  <option key={i} value={t.Nama || (t as any).nama} />
                ))}
              </datalist>
            </div>
          </div>
        </div>

        {/* Kronologi */}
        <div className="space-y-2 md:space-y-4 mt-8 md:mt-12">
          <div className="flex flex-col sm:flex-row sm:justify-between sm:items-end mb-2 gap-2">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              KRONOLOGI KEJADIAN
            </label>
            <button
              type="button"
              onClick={() =>
                onAiAnalysis(
                  formData.kronologi || "",
                  formData.kategori_kasus || "",
                )
              }
              disabled={!formData.kronologi || isAiLoading}
              className="text-[10px] md:text-xs font-semibold uppercase tracking-wider bg-indigo-50 text-indigo-600 px-4 md:px-5 py-2 md:py-2.5 rounded-xl flex items-center justify-center gap-2 hover:bg-indigo-100 transition-all disabled:opacity-50 border border-indigo-100 w-full sm:w-auto"
            >
              {isAiLoading ? (
                <Loader2 size={16} className="animate-spin" />
              ) : (
                <BrainCircuit size={16} />
              )}
              Dapatkan Analisis AI
            </button>
          </div>
          <textarea
            rows={6}
            required
            placeholder="Tuliskan urutan kejadian secara lengkap dan objektif..."
            className="w-full p-4 md:p-6 rounded-2xl border border-slate-200 bg-slate-50 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none resize-none text-sm md:text-base leading-relaxed transition-all placeholder:text-slate-400 text-slate-800"
            value={formData.kronologi}
            onChange={(e) =>
              setFormData({ ...formData, kronologi: e.target.value })
            }
          ></textarea>
        </div>

        {aiAnalysis && (
          <div className="p-5 md:p-8 bg-indigo-50/50 rounded-2xl border border-indigo-100 animate-in fade-in slide-in-from-top-4 duration-500 relative mt-6 md:mt-8">
            <button
              type="button"
              onClick={onClearAiAnalysis}
              className="absolute top-4 right-4 md:top-6 md:right-6 text-indigo-400 hover:text-indigo-600 transition-colors"
            >
              <X size={20} className="md:w-6 md:h-6" />
            </button>
            <h4 className="text-lg md:text-xl font-bold text-indigo-900 mb-4 md:mb-6 flex items-center gap-2 md:gap-3">
              <BrainCircuit size={20} className="text-indigo-600 md:w-6 md:h-6" /> Rekomendasi
              Psikologis Gemini AI
            </h4>
            <div className="space-y-4 md:space-y-6 text-xs md:text-sm leading-relaxed text-slate-700">
              <div>
                <span className="font-semibold text-indigo-600 uppercase text-[10px] md:text-[11px] tracking-wider block mb-1">
                  Analisis Kasus:
                </span>{" "}
                {aiAnalysis.analisis}
              </div>
              <div>
                <span className="font-semibold text-indigo-600 uppercase text-[10px] md:text-[11px] tracking-wider block mb-1">
                  Pendekatan Disarankan:
                </span>{" "}
                {aiAnalysis.pendekatan}
              </div>
              <div>
                <span className="font-semibold text-indigo-600 uppercase text-[10px] md:text-[11px] tracking-wider block mb-1">
                  Langkah Strategis:
                </span>
                <ul className="list-disc ml-4 md:ml-6 mt-1 md:mt-2 space-y-1 md:space-y-2">
                  {aiAnalysis.saran?.map((s: string, i: number) => (
                    <li key={i}>{s}</li>
                  ))}
                </ul>
              </div>
            </div>
          </div>
        )}

        <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 md:gap-x-8 gap-y-6 md:gap-y-8 mt-8 md:mt-12">
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              TINDAK LANJUT UTAMA
            </label>
            <div className="relative">
              <select
                className={inputClass}
                value={formData.tindak_lanjut}
                onChange={(e) =>
                  setFormData({ ...formData, tindak_lanjut: e.target.value })
                }
              >
                {[
                  "Konseling Individu",
                  "Konseling Kelompok",
                  "Panggilan Orang Tua",
                  "Mediasi",
                  "Home Visit",
                  "Skorsing",
                  "Lainnya",
                ].map((k) => (
                  <option key={k} value={k}>
                    {k}
                  </option>
                ))}
              </select>
              <ChevronDown
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-300 pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
          </div>
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              STATUS LAPORAN
            </label>
            <div className="flex flex-col sm:flex-row gap-2 p-1.5 bg-slate-100/50 rounded-xl border border-slate-200">
              {[CaseStatus.BARU, CaseStatus.PROSES, CaseStatus.SELESAI].map(
                (s) => (
                  <button
                    key={s}
                    type="button"
                    onClick={() => setFormData({ ...formData, status: s })}
                    className={`flex-1 py-2.5 md:py-3 rounded-lg text-[10px] md:text-xs font-semibold uppercase tracking-wider transition-all ${formData.status === s ? "bg-white shadow-sm text-indigo-600" : "text-slate-500 hover:text-slate-700 hover:bg-slate-50"}`}
                  >
                    {s}
                  </button>
                ),
              )}
            </div>
          </div>
        </div>

        <div className="p-5 md:p-8 bg-slate-50/50 rounded-2xl border border-dashed border-slate-300 group hover:border-indigo-300 transition-all mt-8 md:mt-12">
          <label className="block text-[10px] md:text-xs font-semibold mb-2 md:mb-4 uppercase text-slate-500 tracking-wider ml-1">
            LAMPIRAN BUKTI FISIK
          </label>
          <input
            type="file"
            multiple
            onChange={handleFileChange}
            className="block w-full text-xs md:text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 md:file:py-2.5 md:file:px-6 file:rounded-xl file:border-0 file:bg-white file:text-indigo-600 file:font-semibold file:text-[10px] md:file:text-xs file:uppercase file:tracking-wider file:shadow-sm hover:file:bg-indigo-50 transition-all cursor-pointer"
          />
          <div className="flex flex-wrap gap-3 md:gap-4 mt-4 md:mt-6">
            {formData.lampiran?.map((f, i) => (
              <div
                key={i}
                className="bg-white px-3 py-2 md:px-4 md:py-3 rounded-xl text-[10px] md:text-xs font-medium flex gap-2 md:gap-3 items-center border border-slate-200 shadow-sm animate-in fade-in zoom-in-90 group-hover:bg-slate-50 transition-all"
              >
                <FileText size={16} className="text-indigo-500 w-4 h-4" />
                <span className="truncate max-w-[100px] md:max-w-[150px] text-slate-700">
                  {f.name}
                </span>
                <button
                  type="button"
                  onClick={() => removeFile(i)}
                  className="text-red-400 hover:text-red-600 transition-colors ml-1 md:ml-2"
                >
                  <X size={16} className="w-4 h-4" />
                </button>
              </div>
            ))}
            {(!formData.lampiran || formData.lampiran.length === 0) && (
              <p className="text-slate-400 text-[10px] md:text-xs italic py-1 ml-1">
                Belum ada file yang diunggah.
              </p>
            )}
          </div>
        </div>

        <div className="pt-8 md:pt-12 border-t border-slate-100 flex flex-col-reverse sm:flex-row justify-end gap-3 md:gap-6 mt-8 md:mt-12">
          <button
            type="button"
            onClick={onCancel}
            className="bg-slate-100 text-slate-600 px-6 py-3 md:px-10 md:py-4 rounded-xl font-semibold hover:bg-slate-200 transition-all text-xs md:text-sm w-full sm:w-auto"
          >
            Batal
          </button>
          <button
            type="submit"
            className={`px-8 py-3 md:px-12 md:py-4 rounded-xl font-bold text-xs md:text-sm flex items-center justify-center gap-2 md:gap-3 transition-all shadow-md ${editingCase ? "bg-amber-500 hover:bg-amber-600 shadow-amber-200" : "bg-indigo-600 hover:bg-indigo-700 shadow-indigo-200"} text-white w-full sm:w-auto`}
          >
            {editingCase ? <Edit size={18} className="w-4 h-4 md:w-5 md:h-5" /> : <CheckCircle2 size={18} className="w-4 h-4 md:w-5 md:h-5" />}{" "}
            {editingCase ? "Update Laporan" : "Simpan Laporan"}
          </button>
        </div>
      </form>
    </div>
  );
};

const ReportView: React.FC<{
  dbState: DatabaseState;
  onEdit: (r: CaseRecord) => void;
  onDelete: (id: number) => void;
  onStatusUpdate: (id: number, s: CaseStatus) => void;
}> = ({ dbState, onEdit, onDelete, onStatusUpdate }) => {
  const [search, setSearch] = useState("");
  const [filterKategori, setFilterKategori] = useState("");
  const [filterStatus, setFilterStatus] = useState("");

  const filtered = useMemo(() => {
    return dbState.kasus
      .filter(
        (k) =>
          (k.nama_siswa.toLowerCase().includes(search.toLowerCase()) ||
            k.kelas.toLowerCase().includes(search.toLowerCase())) &&
          (filterKategori ? k.kategori_kasus === filterKategori : true) &&
          (filterStatus ? k.status === filterStatus : true),
      )
      .sort((a, b) => b.created_at - a.created_at);
  }, [dbState.kasus, search, filterKategori, filterStatus]);

  const handleExport = () => {
    if (filtered.length === 0) return alert("Tidak ada data untuk diekspor");
    const ws = XLSX.utils.json_to_sheet(
      filtered.map(({ lampiran, ...rest }) => ({
        ...rest,
        jum_file: lampiran.length,
      })),
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap_Kasus");
    XLSX.writeFile(
      wb,
      `Laporan_BK_Siswa_${new Date().toISOString().split("T")[0]}.xlsx`,
    );
  };

  return (
    <div className="glass-card p-4 md:p-6 lg:p-10 animate-in fade-in duration-500">
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center mb-6 md:mb-8 gap-4 md:gap-6">
        <div>
          <h3 className="text-xl md:text-2xl font-bold text-slate-900">
            Rekapitulasi Kasus
          </h3>
          <p className="text-xs md:text-sm text-slate-500 mt-1">
            Database lengkap penanganan bimbingan konseling siswa.
          </p>
        </div>
        <div className="flex flex-col sm:flex-row items-center gap-3 w-full lg:w-auto">
          <div className="relative w-full sm:flex-1 lg:min-w-[250px]">
            <Search
              size={18}
              className="absolute left-3 md:left-4 top-1/2 -translate-y-1/2 text-slate-400 w-4 h-4 md:w-5 md:h-5"
            />
            <input
              type="text"
              placeholder="Cari Siswa/Kelas..."
              className="w-full pl-10 md:pl-12 pr-4 py-2 md:py-2.5 rounded-xl bg-slate-50 border border-slate-200 focus:ring-2 focus:ring-indigo-500 outline-none text-xs md:text-sm transition-all"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
          <button
            onClick={handleExport}
            className="bg-emerald-600 text-white px-4 md:px-5 py-2 md:py-2.5 rounded-xl text-xs md:text-sm font-semibold hover:bg-emerald-700 transition flex items-center justify-center gap-2 shadow-sm shadow-emerald-100 w-full sm:w-auto"
          >
            <Download size={16} className="w-4 h-4" /> Excel
          </button>
        </div>
      </div>

      <div className="flex flex-col sm:flex-row flex-wrap gap-2 md:gap-3 mb-4 md:mb-6">
        <select
          className="bg-slate-50 px-3 md:px-4 py-2 rounded-xl text-xs md:text-sm font-medium border border-slate-200 outline-none text-slate-700 focus:ring-2 focus:ring-indigo-500 transition-all w-full sm:w-auto"
          value={filterKategori}
          onChange={(e) => setFilterKategori(e.target.value)}
        >
          <option value="">Semua Kategori</option>
          {[
            "Kedisiplinan",
            "Etika",
            "Akademik",
            "Bullying",
            "Perkelahian",
            "Merokok",
            "Narkoba",
            "Lainnya",
          ].map((k) => (
            <option key={k} value={k}>
              {k}
            </option>
          ))}
        </select>
        <select
          className="bg-slate-50 px-3 md:px-4 py-2 rounded-xl text-xs md:text-sm font-medium border border-slate-200 outline-none text-slate-700 focus:ring-2 focus:ring-indigo-500 transition-all w-full sm:w-auto"
          value={filterStatus}
          onChange={(e) => setFilterStatus(e.target.value)}
        >
          <option value="">Semua Status</option>
          {[CaseStatus.BARU, CaseStatus.PROSES, CaseStatus.SELESAI].map((s) => (
            <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>
      </div>

      <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white">
        <table className="w-full text-left border-collapse min-w-[800px] md:min-w-[1000px]">
          <thead>
            <tr className="text-[10px] md:text-xs uppercase font-semibold text-slate-500 border-b border-slate-200 bg-slate-50">
              <th className="p-3 md:p-4">Tanggal</th>
              <th className="p-3 md:p-4">Informasi Siswa</th>
              <th className="p-3 md:p-4">Kategori</th>
              <th className="p-3 md:p-4">Tindak Lanjut</th>
              <th className="p-3 md:p-4 text-center">Berkas</th>
              <th className="p-3 md:p-4 text-center">Status</th>
              <th className="p-3 md:p-4 text-center">Aksi</th>
            </tr>
          </thead>
          <tbody className="text-xs md:text-sm divide-y divide-slate-100">
            {filtered.length === 0 ? (
              <tr>
                <td colSpan={7} className="p-16 text-center">
                  <div className="flex flex-col items-center opacity-40 grayscale">
                    <Database size={48} className="mb-3 text-slate-400" />
                    <p className="text-slate-500 font-medium text-sm">
                      Belum ada data ditemukan
                    </p>
                  </div>
                </td>
              </tr>
            ) : (
              filtered.map((k) => (
                <tr
                  key={k.id}
                  className="hover:bg-slate-50 transition-colors group"
                >
                  <td className="p-3 md:p-4 whitespace-nowrap text-slate-600">
                    {k.tanggal}
                  </td>
                  <td className="p-3 md:p-4">
                    <div className="font-semibold text-slate-900">
                      {k.nama_siswa}
                    </div>
                    <div className="text-[10px] md:text-xs text-slate-500 mt-0.5">
                      Kelas {k.kelas}
                    </div>
                  </td>
                  <td className="p-3 md:p-4">
                    <span className="bg-indigo-50 text-indigo-700 px-2 md:px-2.5 py-1 rounded-md text-[10px] md:text-xs font-medium">
                      {k.kategori_kasus}
                    </span>
                  </td>
                  <td className="p-3 md:p-4">
                    <div
                      className="max-w-[150px] md:max-w-[200px] truncate font-medium text-slate-700"
                      title={k.tindak_lanjut}
                    >
                      {k.tindak_lanjut}
                    </div>
                    <div className="text-[10px] md:text-xs text-slate-500 mt-0.5">
                      Oleh: {k.guru_bk}
                    </div>
                  </td>
                  <td className="p-3 md:p-4 text-center">
                    {k.lampiran.length > 0 ? (
                      <button
                        onClick={() =>
                          alert(`Total ${k.lampiran.length} file terlampir`)
                        }
                        className="bg-slate-100 text-slate-600 px-2 md:px-2.5 py-1 rounded-md text-[10px] md:text-xs font-medium hover:bg-slate-200 transition-all flex items-center gap-1 md:gap-1.5 mx-auto"
                      >
                        <FileText size={12} className="md:w-[14px] md:h-[14px]" /> {k.lampiran.length}
                      </button>
                    ) : (
                      <span className="text-slate-300">-</span>
                    )}
                  </td>
                  <td className="p-3 md:p-4 text-center">
                    <select
                      value={k.status}
                      onChange={(e) =>
                        onStatusUpdate(k.id!, e.target.value as CaseStatus)
                      }
                      className={`text-[10px] md:text-xs font-semibold px-2 md:px-2.5 py-1 md:py-1.5 rounded-md border-none outline-none appearance-none text-center cursor-pointer transition-colors ${
                        k.status === CaseStatus.BARU
                          ? "bg-blue-50 text-blue-700"
                          : k.status === CaseStatus.PROSES
                            ? "bg-amber-50 text-amber-700"
                            : "bg-emerald-50 text-emerald-700"
                      }`}
                    >
                      <option value={CaseStatus.BARU}>Baru</option>
                      <option value={CaseStatus.PROSES}>Proses</option>
                      <option value={CaseStatus.SELESAI}>Selesai</option>
                    </select>
                  </td>
                  <td className="p-3 md:p-4 text-center">
                    <div className="flex justify-center gap-1 md:gap-2 opacity-100 md:opacity-0 md:group-hover:opacity-100 transition-opacity">
                      <button
                        onClick={() => onEdit(k)}
                        className="p-1 md:p-1.5 bg-white border border-slate-200 text-slate-600 rounded-lg hover:bg-amber-50 hover:text-amber-600 hover:border-amber-200 transition-colors shadow-sm"
                      >
                        <Edit size={14} className="md:w-4 md:h-4" />
                      </button>
                      <button
                        onClick={() => onDelete(k.id!)}
                        className="p-1 md:p-1.5 bg-white border border-slate-200 text-slate-600 rounded-lg hover:bg-red-50 hover:text-red-600 hover:border-red-200 transition-colors shadow-sm"
                      >
                        <Trash2 size={14} className="md:w-4 md:h-4" />
                      </button>
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default App;
