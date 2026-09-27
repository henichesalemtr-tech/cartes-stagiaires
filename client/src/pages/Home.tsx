import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import QRCode from "qrcode";
import {
  Archive,
  ArrowDownToLine,
  ArrowUpFromLine,
  BadgeCheck,
  BarChart3,
  BookOpen,
  Check,
  ChevronLeft,
  ChevronRight,
  ClipboardList,
  Download,
  FileSpreadsheet,
  FileText,
  ImagePlus,
  Languages,
  LayoutDashboard,
  Menu,
  Pencil,
  Plus,
  Printer,
  QrCode,
  RotateCcw,
  Save,
  Search,
  Settings2,
  ShieldCheck,
  Trash2,
  Upload,
  UserRound,
  UsersRound,
  X,
} from "lucide-react";
import { toast } from "sonner";

type View = "dashboard" | "trainees" | "import" | "design" | "print" | "settings";
type Lang = "ar" | "fr";

type Trainee = {
  id: string;
  firstName: string;
  lastName: string;
  registration: string;
  passport: string;
  specialty: string;
  level: string;
  photo?: string;
};

type Settings = {
  orgAr: string;
  orgFr: string;
  line2Ar: string;
  line2Fr: string;
  line3Ar: string;
  line3Fr: string;
  ministryAr: string;
  ministryFr: string;
  cardLanguage: Lang;
  showFlags: boolean;
  showMinistry: boolean;
  accent: "green" | "blue" | "amber";
};

const STORAGE_KEY = "hakani-cards-state-v1";

const defaultSettings: Settings = {
  orgAr: "المعهد الوطني المتخصص في التكوين المهني",
  orgFr: "Institut National Spécialisé de Formation Professionnelle",
  line2Ar: "حساني عبد الكريم",
  line2Fr: "Hassani Abdelkrim",
  line3Ar: "ولاية أدرار • الجزائر",
  line3Fr: "Adrar • Algérie",
  ministryAr: "وزارة التكوين والتعليم المهنيين",
  ministryFr: "Ministère de la Formation et de l'Enseignement Professionnels",
  cardLanguage: "ar",
  showFlags: true,
  showMinistry: true,
  accent: "green",
};

const demoTrainees: Trainee[] = [
  { id: "tr-01", firstName: "ياسين", lastName: "بن عيسى", registration: "2025/00147", passport: "AA384921", specialty: "الإعلام الآلي", level: "تقني سامي" },
  { id: "tr-02", firstName: "مريم", lastName: "قاسمي", registration: "2025/00148", passport: "AA384922", specialty: "المحاسبة والتسيير", level: "تقني سامي" },
  { id: "tr-03", firstName: "سفيان", lastName: "حسني", registration: "2025/00149", passport: "AA384923", specialty: "الكهرباء الصناعية", level: "تقني" },
  { id: "tr-04", firstName: "أميرة", lastName: "بوشارب", registration: "2025/00150", passport: "AA384924", specialty: "التسويق الرقمي", level: "تقني سامي" },
  { id: "tr-05", firstName: "عبد الرؤوف", lastName: "حداد", registration: "2025/00151", passport: "AA384925", specialty: "الطاقات المتجددة", level: "تقني" },
];

function uid() {
  return `tr-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { trainees: demoTrainees, settings: defaultSettings };
    const parsed = JSON.parse(raw);
    return {
      trainees: Array.isArray(parsed.trainees) ? parsed.trainees : demoTrainees,
      settings: { ...defaultSettings, ...(parsed.settings || {}) },
    };
  } catch {
    return { trainees: demoTrainees, settings: defaultSettings };
  }
}

function saveState(trainees: Trainee[], settings: Settings) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ trainees, settings }));
}

function downloadBlob(content: string, filename: string, type = "application/json") {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function Flags() {
  return (
    <span className="flags" aria-label="الجزائر والنيجر">
      <span className="flag flag-dz"><i /></span>
      <span className="flag flag-ne"><i /></span>
    </span>
  );
}

function MinistryMark({ small = false }: { small?: boolean }) {
  return (
    <span className={`ministry-mark ${small ? "ministry-mark-small" : ""}`} aria-label="شعار وزارة التكوين والتعليم المهنيين">
      <svg viewBox="0 0 64 64" role="img">
        <path d="M32 4 53 13v17c0 14-8.5 24.8-21 30C19.5 54.8 11 44 11 30V13L32 4Z" fill="none" stroke="currentColor" strokeWidth="2.4" />
        <path d="m18 26 14-8 14 8-14 8-14-8Zm6 4v8c4.5 3 11.5 3 16 0v-8" fill="none" stroke="currentColor" strokeWidth="2.3" strokeLinejoin="round" />
        <path d="M51 25v13" stroke="currentColor" strokeWidth="2.3" strokeLinecap="round" />
        <circle cx="51" cy="41.5" r="2" fill="currentColor" />
      </svg>
    </span>
  );
}

function QRImage({ trainee, className = "" }: { trainee: Trainee; className?: string }) {
  const [src, setSrc] = useState("");
  useEffect(() => {
    let alive = true;
    QRCode.toDataURL(JSON.stringify({ registration: trainee.registration, fullName: `${trainee.firstName} ${trainee.lastName}`, passport: trainee.passport }), {
      width: 180,
      margin: 0,
      color: { dark: "#112a20", light: "#ffffff" },
      errorCorrectionLevel: "M",
    }).then((url) => alive && setSrc(url));
    return () => { alive = false; };
  }, [trainee]);
  return src ? <img className={className} src={src} alt="QR" /> : <span className={`qr-placeholder ${className}`}><QrCode size={22} /></span>;
}

function CardPreview({ trainee, settings, compact = false, language, onEdit }: { trainee: Trainee; settings: Settings; compact?: boolean; language?: Lang; onEdit?: () => void }) {
  const lang = language || settings.cardLanguage;
  const isFr = lang === "fr";
  return (
    <div className={`id-card ${compact ? "id-card-compact" : ""} accent-${settings.accent}`} dir={isFr ? "ltr" : "rtl"}>
      <div className="card-wave card-wave-one" />
      <div className="card-wave card-wave-two" />
      <div className="zellige zellige-tl" /><div className="zellige zellige-br" />
      <div className="card-crescent">☪</div>
      <div className="card-header">
        {settings.showFlags && <Flags />}
        <div className="card-titles">
          <span>{isFr ? settings.orgFr : settings.orgAr}</span>
          <strong>{isFr ? settings.line2Fr : settings.line2Ar}</strong>
          <small>{isFr ? settings.line3Fr : settings.line3Ar}</small>
        </div>
        {settings.showMinistry && <MinistryMark small />}
      </div>
      <div className="card-rule" />
      <div className="card-body">
        <div className="photo-box">
          {trainee.photo ? <img src={trainee.photo} alt="صورة المتربص" /> : <UserRound size={compact ? 21 : 28} strokeWidth={1.3} />}
          <span>{isFr ? "PHOTO" : "الصورة"}</span>
        </div>
        <div className="trainee-info">
          <div className="trainee-name">{trainee.firstName} {trainee.lastName}</div>
          <div className="info-row"><span>{isFr ? "Matricule" : "رقم التسجيل"}</span><b>{trainee.registration || "—"}</b></div>
          <div className="info-row"><span>{isFr ? "Passeport" : "جواز السفر"}</span><b>{trainee.passport || "—"}</b></div>
          {!compact && <div className="info-row"><span>{isFr ? "Spécialité" : "التخصص"}</span><b>{trainee.specialty || "—"}</b></div>}
        </div>
        <div className="qr-box"><QRImage trainee={trainee} /><span>{isFr ? "SCAN" : "مسح"}</span></div>
      </div>
      <div className="card-footer">
        <span>{isFr ? settings.ministryFr : settings.ministryAr}</span>
        <span className="valid-chip">2025 — 2026</span>
      </div>
      {onEdit && <button className="card-edit" onClick={onEdit} aria-label="تعديل المتربص"><Pencil size={13} /></button>}
    </div>
  );
}

function Sidebar({ activeView, setActiveView, traineesCount, onClose }: { activeView: View; setActiveView: (view: View) => void; traineesCount: number; onClose?: () => void }) {
  const items: { id: View; label: string; icon: typeof LayoutDashboard; badge?: string }[] = [
    { id: "dashboard", label: "نظرة عامة", icon: LayoutDashboard },
    { id: "trainees", label: "المتربصون", icon: UsersRound, badge: String(traineesCount) },
    { id: "import", label: "استيراد البيانات", icon: FileSpreadsheet },
    { id: "design", label: "تصميم البطاقة", icon: BadgeCheck },
    { id: "print", label: "الطباعة والتصدير", icon: Printer },
  ];
  return (
    <aside className="sidebar">
      <div className="brand-lockup">
        <div className="brand-icon"><MinistryMark /></div>
        <div><strong>بطاقتي</strong><span>نظام بطاقات المتربصين</span></div>
      </div>
      <div className="sidebar-label">مساحة العمل</div>
      <nav>
        {items.map(({ id, label, icon: Icon, badge }) => (
          <button key={id} className={`nav-item ${activeView === id ? "active" : ""}`} onClick={() => { setActiveView(id); onClose?.(); }}>
            <Icon size={18} /><span>{label}</span>{badge && <em>{badge}</em>}
          </button>
        ))}
      </nav>
      <div className="sidebar-bottom">
        <div className="storage-card"><div className="storage-icon"><Archive size={16} /></div><div><strong>تخزين محلي</strong><span>بياناتك على جهازك</span></div><Check size={15} /></div>
        <button className={`nav-item ${activeView === "settings" ? "active" : ""}`} onClick={() => { setActiveView("settings"); onClose?.(); }}><Settings2 size={18} /><span>الإعدادات</span></button>
        <div className="user-mini"><div className="avatar">م</div><div><strong>مسؤول المعهد</strong><span>حساب محلي</span></div><ChevronLeft size={16} /></div>
      </div>
    </aside>
  );
}

function Topbar({ activeView, lang, setLang, onMenu }: { activeView: View; lang: Lang; setLang: (lang: Lang) => void; onMenu: () => void }) {
  const titles: Record<View, [string, string]> = {
    dashboard: ["نظرة عامة", "ملخص نشاط البطاقات"],
    trainees: ["المتربصون", "إدارة السجل وإضافة الصور"],
    import: ["استيراد البيانات", "إضافة دفعة من ملف Excel"],
    design: ["تصميم البطاقة", "خصّص هوية البطاقة ومحتواها"],
    print: ["الطباعة والتصدير", "جهّز بطاقاتك للطباعة على A4"],
    settings: ["الإعدادات", "النسخ الاحتياطي والتهيئة المحلية"],
  };
  const [ar, sub] = titles[activeView];
  return <header className="topbar"><button className="mobile-menu" onClick={onMenu}><Menu size={20} /></button><div><h1>{ar}</h1><p>{sub}</p></div><div className="topbar-actions"><div className="online-pill"><span /> يعمل محلياً</div><div className="lang-toggle"><button className={lang === "ar" ? "selected" : ""} onClick={() => setLang("ar")}>ع</button><button className={lang === "fr" ? "selected" : ""} onClick={() => setLang("fr")}>FR</button><Languages size={15} /></div></div></header>;
}

function Dashboard({ trainees, settings, setView, onEdit }: { trainees: Trainee[]; settings: Settings; setView: (v: View) => void; onEdit: (t: Trainee) => void }) {
  const withPhoto = trainees.filter((t) => t.photo).length;
  const stats = [
    { label: "إجمالي المتربصين", value: trainees.length, note: "في السجل المحلي", icon: UsersRound, tone: "green" },
    { label: "بطاقات جاهزة", value: withPhoto, note: `${Math.round((withPhoto / Math.max(trainees.length, 1)) * 100)}% من السجل`, icon: BadgeCheck, tone: "amber" },
    { label: "بانتظار الصور", value: trainees.length - withPhoto, note: "يمكن إضافتها من السجل", icon: ImagePlus, tone: "blue" },
  ];
  const recent = trainees.slice(0, 4);
  return <div className="page-content">
    <section className="hero-panel">
      <div className="hero-copy"><div className="eyebrow"><ShieldCheck size={14} /> مركز إصدار البطاقات</div><h2>كل بطاقة تبدأ من <em>بيانات دقيقة.</em></h2><p>أنشئ بطاقات المتربصين بهوية موحّدة، استورد بياناتك دفعة واحدة، ثم اطبعها في دقائق.</p><div className="hero-actions"><button className="button button-primary" onClick={() => setView("trainees")}><Plus size={17} /> إضافة متربص</button><button className="button button-ghost" onClick={() => setView("import")}><Upload size={16} /> استيراد Excel</button></div></div><div className="hero-visual"><div className="floating-card-main"><CardPreview trainee={trainees[0] || demoTrainees[0]} settings={settings} compact /></div><div className="hero-stamp"><span>CR80</span><small>STANDARD</small></div></div>
    </section>
    <div className="stats-grid">{stats.map(({ label, value, note, icon: Icon, tone }) => <div className="stat-card" key={label}><div className={`stat-icon ${tone}`}><Icon size={19} /></div><div><span>{label}</span><strong>{value}</strong><small>{note}</small></div><BarChart3 className="stat-chart" size={31} /></div>)}</div>
    <div className="dashboard-grid"><section className="surface recent-card"><div className="section-heading"><div><span className="section-kicker">آخر الإضافات</span><h3>المتربصون المضافون حديثاً</h3></div><button className="text-button" onClick={() => setView("trainees")}>عرض الكل <ChevronLeft size={15} /></button></div><div className="recent-list">{recent.map((t, index) => <button className="recent-row" key={t.id} onClick={() => onEdit(t)}><div className="row-avatar">{t.photo ? <img src={t.photo} alt="" /> : t.firstName.charAt(0)}</div><div className="row-name"><strong>{t.firstName} {t.lastName}</strong><span>{t.specialty}</span></div><span className="row-code">{t.registration}</span><span className={`status-dot ${t.photo ? "ready" : "pending"}`}>{t.photo ? "جاهزة" : "صورة ناقصة"}</span><ChevronLeft size={15} className="row-chevron" /></button>)}</div></section><section className="surface quick-card"><div className="section-heading"><div><span className="section-kicker">اختصارات</span><h3>ابدأ من هنا</h3></div></div><button className="quick-action" onClick={() => setView("import")}><div className="quick-icon import"><FileSpreadsheet size={19} /></div><div><strong>استيراد ملف Excel</strong><span>أضف عشرات المتربصين مرة واحدة</span></div><ChevronLeft size={16} /></button><button className="quick-action" onClick={() => setView("design")}><div className="quick-icon design"><BadgeCheck size={19} /></div><div><strong>تخصيص البطاقة</strong><span>عدّل العناوين واللغة والألوان</span></div><ChevronLeft size={16} /></button><button className="quick-action" onClick={() => setView("print")}><div className="quick-icon print"><Printer size={19} /></div><div><strong>تجهيز الطباعة</strong><span>9 بطاقات في ورقة A4</span></div><ChevronLeft size={16} /></button></section></div>
  </div>;
}

function TraineesView({ trainees, setTrainees, settings, openEditor }: { trainees: Trainee[]; setTrainees: (t: Trainee[]) => void; settings: Settings; openEditor: (t?: Trainee) => void }) {
  const [query, setQuery] = useState("");
  const [selected, setSelected] = useState<string[]>([]);
  const filtered = useMemo(() => trainees.filter((t) => `${t.firstName} ${t.lastName} ${t.registration} ${t.passport} ${t.specialty}`.toLowerCase().includes(query.toLowerCase())), [trainees, query]);
  const remove = (id: string) => { if (confirm("حذف هذا المتربص من السجل المحلي؟")) { setTrainees(trainees.filter((t) => t.id !== id)); toast.success("تم حذف المتربص"); } };
  const toggle = (id: string) => setSelected(selected.includes(id) ? selected.filter((x) => x !== id) : [...selected, id]);
  const allVisibleSelected = filtered.length > 0 && filtered.every((t) => selected.includes(t.id));
  return <div className="page-content"><div className="view-toolbar"><div><span className="section-kicker">السجل المحلي</span><h2 className="page-title">قائمة المتربصين <span>{trainees.length}</span></h2></div><div className="toolbar-actions"><div className="search-box"><Search size={17} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="ابحث بالاسم أو رقم التسجيل..." /></div><button className="button button-primary" onClick={() => openEditor()}><Plus size={17} /> متربص جديد</button></div></div><section className="surface table-surface"><div className="table-top"><div className="table-meta">{selected.length ? <span className="selected-count">تم تحديد {selected.length}</span> : <span>السجلات المضافة تظهر هنا</span>}{selected.length > 0 && <button className="link-button" onClick={() => setSelected([])}>إلغاء التحديد</button>}</div><div className="table-tools"><span className="filter-chip"><span className="filter-dot" /> الكل <b>{filtered.length}</b></span></div></div><div className="trainees-table"><div className="table-row table-head"><div><input type="checkbox" checked={allVisibleSelected} onChange={() => setSelected(allVisibleSelected ? selected.filter((id) => !filtered.some((t) => t.id === id)) : Array.from(new Set([...selected, ...filtered.map((t) => t.id)])))} /></div><span>المتربص</span><span>رقم التسجيل</span><span>جواز السفر</span><span>التخصص / المستوى</span><span>حالة البطاقة</span><span></span></div>{filtered.map((t) => <div className="table-row" key={t.id}><div><input type="checkbox" checked={selected.includes(t.id)} onChange={() => toggle(t.id)} /></div><div className="trainee-cell"><div className="table-avatar">{t.photo ? <img src={t.photo} alt="" /> : t.firstName.charAt(0)}</div><div><strong>{t.firstName} {t.lastName}</strong><small>أضيف محلياً</small></div></div><span className="mono-text">{t.registration}</span><span className="mono-text muted">{t.passport || "—"}</span><div className="specialty-cell"><strong>{t.specialty || "—"}</strong><small>{t.level || "—"}</small></div><span className={`table-status ${t.photo ? "ready" : "pending"}`}><i /> {t.photo ? "جاهزة للطباعة" : "تحتاج صورة"}</span><div className="row-actions"><button onClick={() => openEditor(t)} aria-label="تعديل"><Pencil size={15} /></button><button onClick={() => remove(t.id)} aria-label="حذف"><Trash2 size={15} /></button></div></div>)}{filtered.length === 0 && <div className="empty-state"><UsersRound size={32} /><strong>لا توجد نتائج</strong><span>جرّب كلمة بحث أخرى أو أضف متربصاً جديداً</span></div>}</div></section></div>;
}

function ImportView({ onImport, onExport, trainees }: { onImport: (file: File) => void; onExport: () => void; trainees: Trainee[] }) {
  const inputRef = useRef<HTMLInputElement>(null);
  const jsonRef = useRef<HTMLInputElement>(null);
  const downloadTemplate = () => downloadBlob("الاسم,اللقب,رقم التسجيل,رقم جواز السفر,التخصص,المستوى\nياسين,بن عيسى,2025/00147,AA384921,الإعلام الآلي,تقني سامي\n", "قالب-المتربصين.csv", "text/csv;charset=utf-8");
  return <div className="page-content"><div className="view-toolbar"><div><span className="section-kicker">استيراد آمن</span><h2 className="page-title">أضف بياناتك دفعة واحدة</h2><p className="page-lead">يدعم التطبيق ملفات Excel وCSV، وتبقى البيانات محفوظة محلياً على جهازك.</p></div><button className="button button-ghost" onClick={onExport}><ArrowDownToLine size={16} /> تصدير نسخة احتياطية</button></div><div className="import-layout"><section className="surface import-drop" onClick={() => inputRef.current?.click()}><input ref={inputRef} type="file" hidden accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && onImport(e.target.files[0])} /><div className="upload-orb"><Upload size={25} /></div><h3>اسحب ملف Excel هنا</h3><p>أو انقر لاختيار ملف من جهازك</p><span className="file-types"><FileSpreadsheet size={15} /> XLSX / XLS / CSV</span><button className="button button-primary">اختيار الملف <ArrowUpFromLine size={16} /></button></section><section className="surface import-guide"><div className="guide-head"><div className="guide-icon"><ClipboardList size={19} /></div><div><h3>تنسيق الأعمدة</h3><p>استخدم هذه العناوين لتطابق البيانات تلقائياً</p></div></div><div className="columns-list"><span>الاسم <b>مطلوب</b></span><span>اللقب <b>مطلوب</b></span><span>رقم التسجيل <b>مطلوب</b></span><span>رقم جواز السفر <i>اختياري</i></span><span>التخصص <i>اختياري</i></span><span>المستوى <i>اختياري</i></span></div><button className="text-button wide" onClick={downloadTemplate}><Download size={15} /> تحميل قالب جاهز للتعبئة</button></section></div><section className="surface backup-strip"><div className="backup-icon"><Archive size={18} /></div><div><strong>نسخة احتياطية محلية</strong><span>صدّر {trainees.length} سجلاً إلى ملف JSON، واستعده لاحقاً على نفس الجهاز أو جهاز آخر.</span></div><div className="backup-actions"><button className="button button-ghost" onClick={onExport}><Download size={15} /> تصدير JSON</button><input ref={jsonRef} type="file" hidden accept=".json" onChange={(e) => e.target.files?.[0] && onImport(e.target.files[0])} /><button className="button button-ghost" onClick={() => jsonRef.current?.click()}><RotateCcw size={15} /> استعادة JSON</button></div></section></div>;
}

function DesignView({ settings, setSettings, trainee }: { settings: Settings; setSettings: (s: Settings) => void; trainee: Trainee }) {
  const update = (key: keyof Settings, value: string | boolean) => setSettings({ ...settings, [key]: value });
  return <div className="page-content"><div className="view-toolbar"><div><span className="section-kicker">المعاينة المباشرة</span><h2 className="page-title">صمّم بطاقة المعهد</h2><p className="page-lead">كل تعديل يظهر فوراً على نموذج CR80 في الجهة المقابلة.</p></div><div className="design-actions"><button className={`seg-button ${settings.cardLanguage === "ar" ? "active" : ""}`} onClick={() => update("cardLanguage", "ar")}>العربية</button><button className={`seg-button ${settings.cardLanguage === "fr" ? "active" : ""}`} onClick={() => update("cardLanguage", "fr")}>Français</button></div></div><div className="design-layout"><section className="surface design-form"><div className="form-section"><div className="form-section-title"><span>01</span><div><h3>عناوين البطاقة</h3><p>النصوص التي تظهر أعلى منتصف البطاقة</p></div></div><div className="field-grid"><label><span>العنوان الرئيسي بالعربية</span><input value={settings.orgAr} onChange={(e) => update("orgAr", e.target.value)} /></label><label><span>Titre principal en français</span><input dir="ltr" value={settings.orgFr} onChange={(e) => update("orgFr", e.target.value)} /></label><label><span>اسم المعهد بالعربية</span><input value={settings.line2Ar} onChange={(e) => update("line2Ar", e.target.value)} /></label><label><span>Nom de l'institut</span><input dir="ltr" value={settings.line2Fr} onChange={(e) => update("line2Fr", e.target.value)} /></label><label><span>السطر الثالث بالعربية</span><input value={settings.line3Ar} onChange={(e) => update("line3Ar", e.target.value)} /></label><label><span>Troisième ligne</span><input dir="ltr" value={settings.line3Fr} onChange={(e) => update("line3Fr", e.target.value)} /></label></div></div><div className="form-section"><div className="form-section-title"><span>02</span><div><h3>العناصر المرئية</h3><p>تحكم في الشعارات والهوية اللونية</p></div></div><div className="toggle-list"><label className="toggle-row"><span><Flags /><b>علما الجزائر والنيجر</b></span><input type="checkbox" checked={settings.showFlags} onChange={(e) => update("showFlags", e.target.checked)} /><i /></label><label className="toggle-row"><span><MinistryMark small /><b>شعار الوزارة</b></span><input type="checkbox" checked={settings.showMinistry} onChange={(e) => update("showMinistry", e.target.checked)} /><i /></label></div><div className="color-picker"><span>اللون الأساسي</span><div><button className={settings.accent === "green" ? "selected" : ""} onClick={() => update("accent", "green")}><i className="swatch green" />أخضر جزائري</button><button className={settings.accent === "blue" ? "selected" : ""} onClick={() => update("accent", "blue")}><i className="swatch blue" />أزرق رسمي</button><button className={settings.accent === "amber" ? "selected" : ""} onClick={() => update("accent", "amber")}><i className="swatch amber" />أجوري دافئ</button></div></div></div></section><section className="preview-panel"><div className="preview-label"><span><span className="live-dot" /> معاينة مباشرة</span><small>85.6 × 53.98 mm · CR80</small></div><div className="preview-frame"><CardPreview trainee={trainee} settings={settings} /></div><div className="preview-caption"><span>النموذج: {trainee.firstName} {trainee.lastName}</span><button className="text-button" onClick={() => toast.success("تم حفظ إعدادات التصميم محلياً") }><Save size={15} /> حفظ التصميم</button></div></section></div></div>;
}

function PrintView({ trainees, settings }: { trainees: Trainee[]; settings: Settings }) {
  const [selected, setSelected] = useState(trainees.slice(0, 9).map((t) => t.id));
  useEffect(() => { setSelected(trainees.slice(0, 9).map((t) => t.id)); }, [trainees.length]);
  const chosen = trainees.filter((t) => selected.includes(t.id));
  const toggle = (id: string) => setSelected(selected.includes(id) ? selected.filter((x) => x !== id) : [...selected, id]);
  return <div className="page-content print-workspace"><div className="view-toolbar"><div><span className="section-kicker">جاهز للطباعة</span><h2 className="page-title">بطاقات على ورقة A4</h2><p className="page-lead">حدد البطاقات، ثم اطبع أو اختر “حفظ بتنسيق PDF” من نافذة الطباعة.</p></div><button className="button button-primary" onClick={() => window.print()}><Printer size={17} /> طباعة / تصدير PDF</button></div><div className="print-layout"><section className="surface print-list"><div className="print-list-head"><div><strong>اختيار البطاقات</strong><span>{selected.length} من {trainees.length} محددة</span></div><button className="link-button" onClick={() => setSelected(selected.length === trainees.length ? [] : trainees.map((t) => t.id))}>{selected.length === trainees.length ? "إلغاء الكل" : "تحديد الكل"}</button></div>{trainees.map((t) => <label className={`print-select-row ${selected.includes(t.id) ? "selected" : ""}`} key={t.id}><input type="checkbox" checked={selected.includes(t.id)} onChange={() => toggle(t.id)} /><div className="mini-print-avatar">{t.photo ? <img src={t.photo} alt="" /> : t.firstName.charAt(0)}</div><div><strong>{t.firstName} {t.lastName}</strong><span>{t.registration}</span></div><BadgeCheck size={15} className={t.photo ? "text-green" : "text-muted"} /></label>)}</section><section className="print-preview-wrap"><div className="print-preview-toolbar"><span>ورقة A4 · شبكة 3 × 3</span><span className="scale-note">نسبة CR80 محفوظة</span></div><div className="a4-sheet"><div className="print-grid">{chosen.slice(0, 9).map((t) => <CardPreview key={t.id} trainee={t} settings={settings} compact />)}{Array.from({ length: Math.max(0, 9 - chosen.length) }).map((_, i) => <div key={`empty-${i}`} className="empty-print-cell"><Plus size={16} /></div>)}</div></div></section></div></div>;
}

function SettingsView({ trainees, settings, setTrainees, setSettings }: { trainees: Trainee[]; settings: Settings; setTrainees: (t: Trainee[]) => void; setSettings: (s: Settings) => void }) {
  const reset = () => { if (confirm("إعادة البيانات التجريبية وحذف السجل الحالي؟")) { setTrainees(demoTrainees); setSettings(defaultSettings); toast.success("تمت إعادة التهيئة"); } };
  return <div className="page-content"><div className="view-toolbar"><div><span className="section-kicker">مساحة محلية</span><h2 className="page-title">الإعدادات والنسخ الاحتياطي</h2><p className="page-lead">لا توجد قاعدة بيانات خارجية — كل شيء محفوظ في متصفحك على هذا الجهاز.</p></div></div><div className="settings-grid"><section className="surface settings-card"><div className="settings-icon"><Archive size={20} /></div><h3>التخزين المحلي</h3><p>بيانات المتربصين، صورهم، وإعدادات البطاقة محفوظة داخل هذا المتصفح باستخدام localStorage.</p><div className="setting-status"><span className="status-check"><Check size={14} /></span><div><strong>تخزين نشط</strong><small>{trainees.length} سجلاً محفوظاً</small></div></div></section><section className="surface settings-card"><div className="settings-icon"><FileText size={20} /></div><h3>ملف المشروع</h3><p>أنشئ نسخة JSON احتياطية بانتظام، خاصة قبل نقل المشروع إلى حاسوب آخر أو تنظيف بيانات المتصفح.</p><div className="settings-note"><ShieldCheck size={16} /><span>النسخ الاحتياطي لا يحتوي على أي بيانات خارج جهازك.</span></div></section><section className="surface settings-card danger-card"><div className="settings-icon"><RotateCcw size={20} /></div><h3>إعادة التهيئة</h3><p>إرجاع التطبيق إلى بيانات العرض التجريبية وإعدادات البطاقة الافتراضية.</p><button className="button button-danger" onClick={reset}><RotateCcw size={15} /> إعادة البيانات التجريبية</button></section></div></div>;
}

function TraineeEditor({ trainee, onSave, onClose }: { trainee?: Trainee; onSave: (t: Trainee) => void; onClose: () => void }) {
  const [form, setForm] = useState<Trainee>(trainee || { id: uid(), firstName: "", lastName: "", registration: "", passport: "", specialty: "", level: "" });
  const imageRef = useRef<HTMLInputElement>(null);
  const set = (key: keyof Trainee, value: string) => setForm({ ...form, [key]: value });
  const chooseImage = (file?: File) => { if (!file) return; const reader = new FileReader(); reader.onload = () => set("photo", String(reader.result)); reader.readAsDataURL(file); };
  const submit = (e: React.FormEvent) => { e.preventDefault(); if (!form.firstName || !form.lastName || !form.registration) { toast.error("أدخل الاسم واللقب ورقم التسجيل"); return; } onSave(form); };
  return <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}><form className="editor-modal" onSubmit={submit}><div className="modal-head"><div><span className="section-kicker">بطاقة متربص</span><h2>{trainee ? "تعديل البيانات" : "إضافة متربص جديد"}</h2></div><button type="button" className="icon-button" onClick={onClose}><X size={18} /></button></div><div className="editor-body"><div className="photo-upload"><div className="editor-photo">{form.photo ? <img src={form.photo} alt="" /> : <UserRound size={27} />}</div><input ref={imageRef} type="file" hidden accept="image/*" onChange={(e) => chooseImage(e.target.files?.[0])} /><button type="button" className="text-button" onClick={() => imageRef.current?.click()}><ImagePlus size={15} /> {form.photo ? "تغيير الصورة" : "إضافة صورة"}</button><small>JPG أو PNG · صورة شخصية</small></div><div className="editor-fields"><label><span>الاسم <b>*</b></span><input autoFocus value={form.firstName} onChange={(e) => set("firstName", e.target.value)} placeholder="مثال: ياسين" /></label><label><span>اللقب <b>*</b></span><input value={form.lastName} onChange={(e) => set("lastName", e.target.value)} placeholder="مثال: بن عيسى" /></label><label><span>رقم التسجيل <b>*</b></span><input value={form.registration} onChange={(e) => set("registration", e.target.value)} placeholder="2025/00147" /></label><label><span>رقم جواز السفر</span><input dir="ltr" value={form.passport} onChange={(e) => set("passport", e.target.value)} placeholder="AA384921" /></label><label><span>التخصص</span><input value={form.specialty} onChange={(e) => set("specialty", e.target.value)} placeholder="الإعلام الآلي" /></label><label><span>المستوى</span><input value={form.level} onChange={(e) => set("level", e.target.value)} placeholder="تقني سامي" /></label></div></div><div className="modal-foot"><button type="button" className="button button-ghost" onClick={onClose}>إلغاء</button><button type="submit" className="button button-primary"><Save size={16} /> حفظ المتربص</button></div></form></div>;
}

export default function Home() {
  const initial = useMemo(loadState, []);
  const [trainees, setTraineesState] = useState<Trainee[]>(initial.trainees);
  const [settings, setSettingsState] = useState<Settings>(initial.settings);
  const [activeView, setActiveView] = useState<View>("dashboard");
  const [lang, setLang] = useState<Lang>("ar");
  const [editor, setEditor] = useState<{ open: boolean; trainee?: Trainee }>({ open: false });
  const [mobileOpen, setMobileOpen] = useState(false);
  const setTrainees = (next: Trainee[]) => { setTraineesState(next); saveState(next, settings); };
  const setSettings = (next: Settings) => { setSettingsState(next); saveState(trainees, next); };
  useEffect(() => saveState(trainees, settings), [trainees, settings]);
  const saveTrainee = (next: Trainee) => { const exists = trainees.some((t) => t.id === next.id); const updated = exists ? trainees.map((t) => t.id === next.id ? next : t) : [next, ...trainees]; setTrainees(updated); setEditor({ open: false }); toast.success(exists ? "تم تحديث بيانات المتربص" : "تمت إضافة المتربص"); };
  const importFile = async (file: File) => {
    try {
      if (file.name.endsWith(".json")) {
        const parsed = JSON.parse(await file.text());
        const nextTrainees = Array.isArray(parsed) ? parsed : parsed.trainees;
        if (!Array.isArray(nextTrainees)) throw new Error("invalid");
        setTrainees(nextTrainees.map((t: Trainee) => ({ ...t, id: t.id || uid() })));
        if (parsed.settings) setSettings({ ...settings, ...parsed.settings });
        toast.success(`تمت استعادة ${nextTrainees.length} سجلاً`);
        return;
      }
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
      const value = (row: Record<string, unknown>, keys: string[]) => { const key = Object.keys(row).find((k) => keys.some((candidate) => k.trim().toLowerCase().includes(candidate))); return key ? String(row[key] ?? "").trim() : ""; };
      const imported = rows.map((row) => ({ id: uid(), firstName: value(row, ["الاسم", "prenom", "first", "name"]), lastName: value(row, ["اللقب", "nom", "last"]), registration: value(row, ["التسجيل", "matricule", "registration"]), passport: value(row, ["جواز", "passeport", "passport"]), specialty: value(row, ["التخصص", "specialite", "spécialité", "specialty"]), level: value(row, ["المستوى", "niveau", "level"]) })).filter((t) => t.firstName || t.lastName || t.registration);
      if (!imported.length) throw new Error("empty");
      setTrainees([...imported, ...trainees]);
      toast.success(`تم استيراد ${imported.length} متربصاً بنجاح`);
      setActiveView("trainees");
    } catch { toast.error("تعذر قراءة الملف. تحقق من التنسيق ثم حاول مجدداً."); }
  };
  const exportBackup = () => downloadBlob(JSON.stringify({ version: 1, exportedAt: new Date().toISOString(), trainees, settings }, null, 2), `بطاقات-المتربصين-${new Date().toISOString().slice(0, 10)}.json`);
  const currentTrainee = trainees[0] || demoTrainees[0];
  return <div className="app-shell" dir="rtl"><div className={`mobile-overlay ${mobileOpen ? "show" : ""}`} onClick={() => setMobileOpen(false)} /><div className={`sidebar-wrap ${mobileOpen ? "open" : ""}`}><Sidebar activeView={activeView} setActiveView={setActiveView} traineesCount={trainees.length} onClose={() => setMobileOpen(false)} /></div><main className="main-area"><Topbar activeView={activeView} lang={lang} setLang={setLang} onMenu={() => setMobileOpen(true)} />{activeView === "dashboard" && <Dashboard trainees={trainees} settings={settings} setView={setActiveView} onEdit={(t) => setEditor({ open: true, trainee: t })} />}{activeView === "trainees" && <TraineesView trainees={trainees} setTrainees={setTrainees} settings={settings} openEditor={(t) => setEditor({ open: true, trainee: t })} />}{activeView === "import" && <ImportView onImport={importFile} onExport={exportBackup} trainees={trainees} />}{activeView === "design" && <DesignView settings={settings} setSettings={setSettings} trainee={currentTrainee} />}{activeView === "print" && <PrintView trainees={trainees} settings={settings} />}{activeView === "settings" && <SettingsView trainees={trainees} settings={settings} setTrainees={setTrainees} setSettings={setSettings} />}</main>{editor.open && <TraineeEditor trainee={editor.trainee} onSave={saveTrainee} onClose={() => setEditor({ open: false })} />}</div>;
}
