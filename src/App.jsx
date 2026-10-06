import { useEffect, useMemo, useState } from "react";
import { supabase } from "./lib/supabase";
import * as XLSX from "xlsx";
import "./App.css";

const MONEY = new Intl.NumberFormat("en-TZ");

function money(value) {
  return `TZS ${MONEY.format(Number(value || 0))}`;
}

function number(value) {
  return Number(value || 0);
}

function positiveNumber(value) {
  return Math.max(0, number(value));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function todayDateInput() {
  return new Date().toISOString().slice(0, 10);
}

function normalizePhone(value) {
  return String(value || "").replace(/[^0-9+]/g, "");
}

function todayStart() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d.toISOString();
}

function monthStart() {
  const d = new Date();
  d.setDate(1);
  d.setHours(0, 0, 0, 0);
  return d.toISOString();
}

function localDateTimeValue() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(
    d.getDate()
  )}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function formatDate(value) {
  if (!value) return "-";
  return new Date(value).toLocaleString("en-TZ", {
    dateStyle: "medium",
    timeStyle: "short",
  });
}

function exportToExcel(rows, fileName, sheetName = "Data") {
  if (!rows || !rows.length) {
    alert("Hakuna data ya ku-export kwa sasa.");
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.slice(0, 31));
  XLSX.writeFile(workbook, fileName);
}

function exportRowsWithFormattedDates(rows, dateKeys = []) {
  return rows.map((row) => {
    const copy = { ...row };
    dateKeys.forEach((key) => {
      if (copy[key]) copy[key] = formatDate(copy[key]);
    });
    return copy;
  });
}

function App() {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(true);
  const [authMessage, setAuthMessage] = useState("");

  useEffect(() => {
    let mounted = true;

    async function initializeAuth() {
      try {
        const url = new URL(window.location.href);
        const code = url.searchParams.get("code");

        if (code) {
          const { data, error } = await supabase.auth.exchangeCodeForSession(code);

          if (error) {
            console.error("Email confirmation code exchange failed:", error);
            if (mounted) {
              setAuthMessage(
                "Email imethibitishwa, lakini session haikufunguka. Jaribu kuingia tena kwa email na password yako."
              );
            }
          } else if (data?.session && mounted) {
            setSession(data.session);
          }

          url.searchParams.delete("code");
          window.history.replaceState(
            {},
            document.title,
            url.pathname + url.search + url.hash
          );
        }

        const tokenHash = url.searchParams.get("token_hash");
        const tokenType = url.searchParams.get("type");

        if (!code && tokenHash && tokenType === "email") {
          const { data, error } = await supabase.auth.verifyOtp({
            token_hash: tokenHash,
            type: "email",
          });

          if (error) {
            console.error("Email confirmation verification failed:", error);
            if (mounted) {
              setAuthMessage(
                "Link ya confirmation imekwisha au haikukamilika. Tuma confirmation email nyingine."
              );
            }
          } else if (data?.session && mounted) {
            setSession(data.session);
          }

          url.searchParams.delete("token_hash");
          url.searchParams.delete("type");
          window.history.replaceState(
            {},
            document.title,
            url.pathname + url.search + url.hash
          );
        }

        const hashParams = new URLSearchParams(
          window.location.hash.replace(/^#/, "")
        );
        const hashError = hashParams.get("error_description");

        if (hashError && mounted) {
          setAuthMessage(decodeURIComponent(hashError.replace(/\+/g, " ")));
        }

        const { data: sessionData } = await supabase.auth.getSession();

        if (mounted) {
          setSession(sessionData.session);
          setLoading(false);
        }
      } catch (err) {
        console.error("Auth initialization failed:", err);

        if (mounted) {
          setAuthMessage(
            "Kuna tatizo kwenye confirmation ya email. Jaribu tena."
          );
          setLoading(false);
        }
      }
    }

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((_event, newSession) => {
      if (!mounted) return;
      setSession(newSession);
      setLoading(false);
    });

    initializeAuth();

    return () => {
      mounted = false;
      subscription.unsubscribe();
    };
  }, []);

  if (loading) {
    return (
      <div className="loading-screen">
        <div className="loading-box">
          <div className="logo-circle">B</div>
          <h2>Bless Stationery</h2>
          <p>Inafungua mfumo...</p>
          <div className="spinner" />
        </div>
      </div>
    );
  }

  if (!session) {
    return <Login initialMessage={authMessage} />;
  }

  return <System user={session.user} />;
}

const REGISTRATION_PLANS = [
  {
    id: "FREE_TRIAL",
    title: "FREE TRIAL",
    price: "TZS 0",
    period: "Siku 14",
    description: "Jaribu mfumo bila malipo kwa siku 14.",
  },
  {
    id: "MONTHLY",
    title: "MWEZI",
    price: "TZS 10,000",
    period: "Siku 30",
    description: "Mpango wa mwezi kwa biashara yako.",
  },
  {
    id: "YEARLY",
    title: "MWAKA",
    price: "TZS 50,000",
    period: "Mwaka 1",
    description: "Mpango wa mwaka mmoja kwa biashara yako.",
  },
];

const FALLBACK_COUNTRIES = [
  { name: "Tanzania", iso2: "TZ" },
  { name: "Kenya", iso2: "KE" },
  { name: "Uganda", iso2: "UG" },
  { name: "Rwanda", iso2: "RW" },
  { name: "Burundi", iso2: "BI" },
  { name: "Zambia", iso2: "ZM" },
  { name: "Malawi", iso2: "MW" },
  { name: "Mozambique", iso2: "MZ" },
  { name: "South Africa", iso2: "ZA" },
  { name: "United States", iso2: "US" },
  { name: "United Kingdom", iso2: "GB" },
];

async function fetchCountries() {
  try {
    const response = await fetch("https://countriesnow.space/api/v0.1/countries/positions");
    if (!response.ok) throw new Error("countries request failed");
    const json = await response.json();
    const rows = Array.isArray(json?.data) ? json.data : [];
    return rows
      .map((x) => ({ name: x.name, iso2: x.iso2 }))
      .filter((x) => x.name && x.iso2)
      .sort((a, b) => a.name.localeCompare(b.name));
  } catch {
    return FALLBACK_COUNTRIES;
  }
}

const TANZANIA_REGIONS = [
  "Arusha",
  "Dar es Salaam",
  "Dodoma",
  "Geita",
  "Iringa",
  "Kagera",
  "Katavi",
  "Kigoma",
  "Kilimanjaro",
  "Lindi",
  "Manyara",
  "Mara",
  "Mbeya",
  "Morogoro",
  "Mtwara",
  "Mwanza",
  "Njombe",
  "Pemba North",
  "Pemba South",
  "Pwani",
  "Rukwa",
  "Ruvuma",
  "Shinyanga",
  "Simiyu",
  "Singida",
  "Songwe",
  "Tabora",
  "Tanga",
  "Zanzibar North",
  "Zanzibar South",
  "Zanzibar West",
];

function isTanzania(country) {
  const value = String(country || "").trim().toLowerCase();
  return value === "tanzania" || value === "tanzania, united republic of";
}

async function fetchStates(country) {
  if (!country) return [];
  try {
    const response = await fetch("https://countriesnow.space/api/v0.1/countries/states", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ country }),
    });
    if (!response.ok) throw new Error("states request failed");
    const json = await response.json();
    const rows = (json?.data?.states || []).map((x) => x.name).filter(Boolean);
    if (rows.length) return rows;
    if (isTanzania(country)) return TANZANIA_REGIONS;
    return [];
  } catch {
    if (isTanzania(country)) return TANZANIA_REGIONS;
    return [];
  }
}

async function fetchCities(country, state) {
  if (!country || !state) return [];
  try {
    const response = await fetch("https://countriesnow.space/api/v0.1/countries/state/cities", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ country, state }),
    });
    if (!response.ok) throw new Error("cities request failed");
    const json = await response.json();
    return Array.isArray(json?.data) ? json.data.filter(Boolean) : [];
  } catch {
    return [];
  }
}

function Login({ initialMessage = "" }) {
  const [mode, setMode] = useState("login");

  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");

  const [businessName, setBusinessName] = useState("");
  const [fullName, setFullName] = useState("");
  const [phone, setPhone] = useState("");
  const [address, setAddress] = useState("");
  const [country, setCountry] = useState("");
  const [region, setRegion] = useState("");
  const [district, setDistrict] = useState("");
  const [ward, setWard] = useState("");
  const [street, setStreet] = useState("");
  const [selectedPlan, setSelectedPlan] = useState("FREE_TRIAL");
  const [countries, setCountries] = useState(FALLBACK_COUNTRIES);
  const [regions, setRegions] = useState([]);
  const [cities, setCities] = useState([]);
  const [locationLoading, setLocationLoading] = useState(false);
  const [confirmPassword, setConfirmPassword] = useState("");

  const [showPassword, setShowPassword] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [message, setMessage] = useState(initialMessage);

  useEffect(() => {
    if (mode !== "register") return;
    let alive = true;
    fetchCountries().then((rows) => {
      if (alive && rows.length) setCountries(rows);
    });
    return () => {
      alive = false;
    };
  }, [mode]);

  useEffect(() => {
    if (mode !== "register" || !country) {
      setRegions([]);
      return;
    }
    let alive = true;
    setLocationLoading(true);
    setRegion("");
    setDistrict("");
    setCities([]);
    fetchStates(country).then((rows) => {
      if (alive) {
        setRegions(rows);
        setLocationLoading(false);
      }
    });
    return () => {
      alive = false;
    };
  }, [mode, country]);

  useEffect(() => {
    if (mode !== "register" || !country || !region) {
      setCities([]);
      return;
    }
    let alive = true;
    setLocationLoading(true);
    setDistrict("");
    fetchCities(country, region).then((rows) => {
      if (alive) {
        setCities(rows);
        setLocationLoading(false);
      }
    });
    return () => {
      alive = false;
    };
  }, [mode, country, region]);

  function switchMode(nextMode) {
    setMode(nextMode);
    setError("");
    setMessage("");
  }

  async function handleLogin(e) {
    e.preventDefault();
    setError("");
    setMessage("");
    setBusy(true);

    const { error } = await supabase.auth.signInWithPassword({
      email: email.trim(),
      password,
    });

    if (error) {
      const authError = String(error.message || "").toLowerCase();

      if (authError.includes("email not confirmed")) {
        setError(
          "Email yako bado haijathibitishwa. Fungua confirmation email ya Supabase, kisha ujaribu INGIA tena."
        );
      } else if (authError.includes("invalid login credentials")) {
        setError("Email au password sio sahihi.");
      } else {
        setError(error.message || "Login imeshindikana.");
      }
    }

    setBusy(false);
  }

  async function handleRegister(e) {
    e.preventDefault();

    setError("");
    setMessage("");

    if (password.length < 6) {
      setError("Password lazima iwe na angalau characters 6.");
      return;
    }

    if (password !== confirmPassword) {
      setError("Password hazifanani.");
      return;
    }

    if (!businessName.trim()) {
      setError("Weka jina la biashara.");
      return;
    }

    if (!fullName.trim()) {
      setError("Weka jina la mmiliki.");
      return;
    }

    if (!phone.trim()) {
      setError("Weka namba ya simu.");
      return;
    }

    if (!country) {
      setError("Chagua nchi.");
      return;
    }

    if (!region) {
      setError("Chagua region/state.");
      return;
    }

    if (!district.trim()) {
      setError("Weka district/city.");
      return;
    }

    if (!ward.trim()) {
      setError("Weka ward.");
      return;
    }

    if (!street.trim()) {
      setError("Weka street / eneo la biashara.");
      return;
    }

    setBusy(true);

    const { data, error } = await supabase.auth.signUp({
      email: email.trim(),
      password,
      options: {
        emailRedirectTo: window.location.origin,
        data: {
          business_name: businessName.trim(),
          full_name: fullName.trim(),
          phone: phone.trim(),
          address: address.trim(),
          selected_plan: selectedPlan,
          country: country.trim(),
          region: region.trim(),
          district: district.trim(),
          ward: ward.trim(),
          street: street.trim(),
        },
      },
    });

    if (error) {
      setError(error.message);
      setBusy(false);
      return;
    }

    if (data?.session) {
      setMessage("Account imetengenezwa. Inafungua mfumo...");
    } else {
      setMessage(
        "Account imetengenezwa. Kama email confirmation imewashwa, fungua email yako kuthibitisha account."
      );
    }

    setBusy(false);
  }

  return (
    <div className="login-page">
      <div className="login-brand">
        <div className="brand-logo">B</div>
        <h1>Bless Stationery</h1>
        <p>Business Management System</p>

        <div className="brand-features">
          <div>✓ Sales & Receipts</div>
          <div>✓ Stock Management</div>
          <div>✓ Expenses & Deposits</div>
          <div>✓ Daily & Monthly Reports</div>
        </div>
      </div>

      <div className="login-side">
        <div className="login-card">
          <div className="mobile-logo">B</div>

          <div className="login-tabs">
            <button
              type="button"
              className={mode === "login" ? "active" : ""}
              onClick={() => switchMode("login")}
            >
              INGIA
            </button>
            <button
              type="button"
              className={mode === "register" ? "active" : ""}
              onClick={() => switchMode("register")}
            >
              JISAJILI
            </button>
          </div>

          {mode === "login" ? (
            <form onSubmit={handleLogin}>
              <h2>Karibu tena</h2>
              <p className="muted">Ingia kwenye mfumo wa Bless Stationery</p>

              {error && <div className="error-box">{error}</div>}
              {message && <div className="success-box">{message}</div>}

              <label>Email</label>
              <input
                type="email"
                placeholder="Weka email yako"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
              />

              <label>Password</label>
              <div className="password-wrap">
                <input
                  type={showPassword ? "text" : "password"}
                  placeholder="Weka password yako"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  required
                />
                <button
                  type="button"
                  className="show-password"
                  onClick={() => setShowPassword(!showPassword)}
                >
                  {showPassword ? "Ficha" : "Onyesha"}
                </button>
              </div>

              <button className="primary-btn login-btn" disabled={busy}>
                {busy ? "Inaingia..." : "INGIA KWENYE MFUMO"}
              </button>

              <p className="register-switch">
                Huna account?{" "}
                <button type="button" onClick={() => switchMode("register")}>
                  Jisajili hapa
                </button>
              </p>
            </form>
          ) : (
            <form onSubmit={handleRegister}>
              <h2>Fungua Account</h2>
              <p className="muted">Anzisha mfumo wa biashara yako</p>

              {error && <div className="error-box">{error}</div>}
              {message && <div className="success-box">{message}</div>}

              <label>Jina la Biashara</label>
              <input
                type="text"
                placeholder="Mfano: Bless Stationery"
                value={businessName}
                onChange={(e) => setBusinessName(e.target.value)}
                required
              />

              <label>Jina la Mmiliki</label>
              <input
                type="text"
                placeholder="Jina lako kamili"
                value={fullName}
                onChange={(e) => setFullName(e.target.value)}
                required
              />

              <label>Email</label>
              <input
                type="email"
                placeholder="mfano@email.com"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
              />

              <label>Phone</label>
              <input
                type="tel"
                placeholder="07XXXXXXXX"
                value={phone}
                onChange={(e) => setPhone(e.target.value)}
                required
              />

              <div className="registration-section-title">Chagua Mpango</div>
              <div className="plan-grid">
                {REGISTRATION_PLANS.map((plan) => (
                  <button
                    key={plan.id}
                    type="button"
                    className={`plan-card ${selectedPlan === plan.id ? "selected" : ""}`}
                    onClick={() => setSelectedPlan(plan.id)}
                  >
                    <strong>{plan.title}</strong>
                    <span className="plan-price">{plan.price}</span>
                    <span className="plan-period">{plan.period}</span>
                    <small>{plan.description}</small>
                  </button>
                ))}
              </div>

              <div className="registration-section-title">Mahali pa Biashara</div>

              <label>Nchi</label>
              <select value={country} onChange={(e) => setCountry(e.target.value)} required>
                <option value="">Chagua nchi</option>
                {countries.map((item) => (
                  <option value={item.name} key={`${item.iso2}-${item.name}`}>
                    {item.name}
                  </option>
                ))}
              </select>

              <label>Region / State</label>
              <select
                value={region}
                onChange={(e) => setRegion(e.target.value)}
                required
                disabled={!country || locationLoading}
              >
                <option value="">
                  {locationLoading ? "Inapakia..." : "Chagua region / state"}
                </option>
                {regions.map((item) => (
                  <option value={item} key={item}>{item}</option>
                ))}
              </select>

              <label>District / City</label>
              {cities.length > 0 ? (
                <select
                  value={district}
                  onChange={(e) => setDistrict(e.target.value)}
                  required
                  disabled={!region}
                >
                  <option value="">Chagua district / city</option>
                  {cities.map((item) => (
                    <option value={item} key={item}>{item}</option>
                  ))}
                </select>
              ) : (
                <input
                  type="text"
                  placeholder="Mfano: Ilala / Dar es Salaam"
                  value={district}
                  onChange={(e) => setDistrict(e.target.value)}
                  required
                />
              )}

              <label>Ward</label>
              <input
                type="text"
                placeholder="Mfano: Kariakoo"
                value={ward}
                onChange={(e) => setWard(e.target.value)}
                required
              />

              <label>Street / Eneo la Biashara</label>
              <input
                type="text"
                placeholder="Mfano: Msimbazi Street, Plot 12"
                value={street}
                onChange={(e) => setStreet(e.target.value)}
                required
              />

              <label>Address ya Ziada (optional)</label>
              <input
                type="text"
                placeholder="Maelezo ya ziada ya anwani"
                value={address}
                onChange={(e) => setAddress(e.target.value)}
              />

              <label>Password</label>
              <div className="password-wrap">
                <input
                  type={showPassword ? "text" : "password"}
                  placeholder="Weka password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  required
                />
                <button
                  type="button"
                  className="show-password"
                  onClick={() => setShowPassword(!showPassword)}
                >
                  {showPassword ? "Ficha" : "Onyesha"}
                </button>
              </div>

              <label>Confirm Password</label>
              <input
                type={showPassword ? "text" : "password"}
                placeholder="Rudia password"
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                required
              />

              <button className="primary-btn login-btn" disabled={busy}>
                {busy ? "Inatengeneza account..." : "JISAJILI"}
              </button>

              <p className="register-switch">
                Tayari una account?{" "}
                <button type="button" onClick={() => switchMode("login")}>
                  Ingia hapa
                </button>
              </p>
            </form>
          )}

          <p className="login-footer">
            © {new Date().getFullYear()} Bless Stationery
          </p>
        </div>
      </div>
    </div>
  );
}

function System({ user }) {
  const [page, setPage] = useState("dashboard");
  const [mobileMenu, setMobileMenu] = useState(false);
  const [staff, setStaff] = useState(null);
  const [business, setBusiness] = useState(null);
  const [isPlatformAdmin, setIsPlatformAdmin] = useState(false);
  const [platformAdminLoading, setPlatformAdminLoading] = useState(true);
  const [subscriptionAccess, setSubscriptionAccess] = useState(null);
  const [subscriptionLoading, setSubscriptionLoading] = useState(true);
  const [refresh, setRefresh] = useState(0);
  const [theme, setTheme] = useState(() => localStorage.getItem(`bless_theme_${user.id}`) || "ocean");
  const [density, setDensity] = useState(() => localStorage.getItem(`bless_density_${user.id}`) || "comfortable");
  const [language, setLanguage] = useState(() => localStorage.getItem(`bless_language_${user.id}`) || "sw");
  const [availableLanguages, setAvailableLanguages] = useState([]);
  const [showWelcome, setShowWelcome] = useState(() => {
    try {
      return sessionStorage.getItem(`bless_welcome_${user.id}`) !== "seen";
    } catch {
      return true;
    }
  });
  const [showCommand, setShowCommand] = useState(false);
  const [pulse, setPulse] = useState({ sales: 0, expenses: 0, products: 0, debts: 0 });
  const businessId = staff?.business_id || null;

  useEffect(() => {
    if (!businessId) return;
    let cancelled = false;
    async function loadPulse() {
      const [salesRes, expenseRes, productRes, debtRes] = await Promise.all([
        supabase.from("sales").select("sales_total").eq("business_id", businessId),
        supabase.from("expenses").select("amount").eq("business_id", businessId),
        supabase.from("products").select("id").eq("business_id", businessId).eq("active", true),
        supabase.from("credit_transactions").select("balance").eq("business_id", businessId).gt("balance", 0),
      ]);
      if (cancelled) return;
      setPulse({
        sales: (salesRes.data || []).reduce((a, x) => a + number(x.sales_total), 0),
        expenses: (expenseRes.data || []).reduce((a, x) => a + number(x.amount), 0),
        products: (productRes.data || []).length,
        debts: (debtRes.data || []).reduce((a, x) => a + number(x.balance), 0),
      });
    }
    loadPulse();
    return () => { cancelled = true; };
  }, [businessId, refresh]);

  useEffect(() => {
    // Google Translate powers the translation, while our own selector keeps
    // language names clear (e.g. Dutch, not "Kiholanzi") and shows "Lugha".
    const languageNames = {
      af: "Afrikaans", sq: "Albanian", am: "Amharic", ar: "Arabic", hy: "Armenian", az: "Azerbaijani", eu: "Basque", be: "Belarusian", bn: "Bengali", bs: "Bosnian", bg: "Bulgarian", ca: "Catalan", ceb: "Cebuano", ny: "Chichewa",
      "zh-CN": "Chinese (Simplified)", "zh-TW": "Chinese (Traditional)", co: "Corsican", hr: "Croatian", cs: "Czech", da: "Danish", nl: "Dutch", en: "English", eo: "Esperanto", et: "Estonian", tl: "Filipino", fi: "Finnish", fr: "French", fy: "Frisian", gl: "Galician", ka: "Georgian", de: "German", el: "Greek", gu: "Gujarati", ht: "Haitian Creole", ha: "Hausa", haw: "Hawaiian", he: "Hebrew", hi: "Hindi", hmn: "Hmong", hu: "Hungarian", is: "Icelandic", ig: "Igbo", id: "Indonesian", ga: "Irish", it: "Italian", ja: "Japanese", jv: "Javanese", kn: "Kannada", kk: "Kazakh", km: "Khmer", rw: "Kinyarwanda", ko: "Korean", ku: "Kurdish", ky: "Kyrgyz", lo: "Lao", la: "Latin", lv: "Latvian", lt: "Lithuanian", lb: "Luxembourgish", mk: "Macedonian", mg: "Malagasy", ms: "Malay", ml: "Malayalam", mt: "Maltese", mi: "Maori", mr: "Marathi", mn: "Mongolian", my: "Myanmar (Burmese)", ne: "Nepali", no: "Norwegian", or: "Odia", ps: "Pashto", fa: "Persian", pl: "Polish", pt: "Portuguese", pa: "Punjabi", ro: "Romanian", ru: "Russian", sm: "Samoan", gd: "Scots Gaelic", sr: "Serbian", st: "Sesotho", sn: "Shona", sd: "Sindhi", si: "Sinhala", sk: "Slovak", sl: "Slovenian", so: "Somali", es: "Spanish", su: "Sundanese", sw: "Kiswahili", sv: "Swedish", tg: "Tajik", ta: "Tamil", tt: "Tatar", te: "Telugu", th: "Thai", tr: "Turkish", tk: "Turkmen", uk: "Ukrainian", ur: "Urdu", ug: "Uyghur", uz: "Uzbek", vi: "Vietnamese", cy: "Welsh", xh: "Xhosa", yi: "Yiddish", yo: "Yoruba", zu: "Zulu"
    };

    const syncLanguageOptions = () => {
      const select = document.querySelector("#google_translate_element select.goog-te-combo");
      if (!select) return false;
      const options = Array.from(select.options)
        .map((option) => ({ value: option.value, label: languageNames[option.value] || option.textContent.trim() }))
        .filter((option) => option.value);
      const ordered = [
        { value: "sw", label: "Kiswahili" },
        ...options.filter((x) => x.value !== "sw"),
      ];
      const seen = new Set();
      const cleanOptions = ordered.filter((x) => !seen.has(x.value) && seen.add(x.value));
      setAvailableLanguages(cleanOptions);
      if (language && language !== "sw" && select.value !== language) {
        select.value = language;
        select.dispatchEvent(new Event("change", { bubbles: true }));
      }
      return true;
    };

    window.googleTranslateElementInit = () => {
      if (!window.google?.translate || !document.getElementById("google_translate_element")) return;
      try {
        new window.google.translate.TranslateElement({
          pageLanguage: "sw",
          autoDisplay: false,
          multilanguagePage: true,
        }, "google_translate_element");
        let tries = 0;
        const timer = setInterval(() => {
          tries += 1;
          if (syncLanguageOptions() || tries >= 30) clearInterval(timer);
        }, 150);
      } catch (err) {
        console.warn("Google Translate init failed:", err);
      }
    };

    if (!document.getElementById("google-translate-script")) {
      const script = document.createElement("script");
      script.id = "google-translate-script";
      script.src = "https://translate.google.com/translate_a/element.js?cb=googleTranslateElementInit";
      script.async = true;
      document.body.appendChild(script);
    } else if (window.google?.translate) {
      window.googleTranslateElementInit();
    }

    return () => {};
  }, []);

  function changeLanguage(value) {
    setLanguage(value);
    localStorage.setItem(`bless_language_${user.id}`, value);
    const select = document.querySelector("#google_translate_element select.goog-te-combo");
    if (!select) return;
    select.value = value;
    select.dispatchEvent(new Event("change", { bubbles: true }));
  }

  useEffect(() => {
    const onKey = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        setShowCommand((v) => !v);
      }
      if (e.key === "Escape") {
        setShowCommand(false);
        setShowWelcome(false);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  function closeWelcome() {
    setShowWelcome(false);
    try { sessionStorage.setItem(`bless_welcome_${user.id}`, "seen"); } catch {}
  }

  useEffect(() => {
    localStorage.setItem(`bless_theme_${user.id}`, theme);
  }, [theme, user.id]);

  useEffect(() => {
    localStorage.setItem(`bless_density_${user.id}`, density);
  }, [density, user.id]);

  useEffect(() => {
    loadPlatformAdmin();
    loadStaff();
  }, [user.id]);

  async function loadPlatformAdmin() {
    setPlatformAdminLoading(true);

    const { data, error } = await supabase
      .from("platform_admins")
      .select("id, active")
      .eq("user_id", user.id)
      .eq("active", true)
      .maybeSingle();

    if (error) {
      console.error("Failed to check platform admin:", error);
      setIsPlatformAdmin(false);
    } else {
      setIsPlatformAdmin(!!data);
    }

    setPlatformAdminLoading(false);
  }

  async function loadStaff() {
    const { data, error } = await supabase
      .from("staff")
      .select("*, businesses(*)")
      .eq("user_id", user.id)
      .maybeSingle();

    if (error) {
      console.error("Failed to load staff/business:", error);
      setStaff(null);
      setBusiness(null);
      return;
    }

    setStaff(data || null);
    setBusiness(data?.businesses || null);

    if (data?.business_id) {
      await checkSubscriptionAccess(data.business_id);
    } else {
      setSubscriptionAccess(null);
      setSubscriptionLoading(false);
    }
  }

  async function checkSubscriptionAccess(targetBusinessId) {
    setSubscriptionLoading(true);

    const { data, error } = await supabase.rpc(
      "business_subscription_active",
      { target_business_id: targetBusinessId }
    );

    if (error) {
      console.error("Failed to check subscription:", error);
      // Fail closed for normal customer businesses. Platform admins are
      // handled separately and internal businesses are exempt in the DB.
      setSubscriptionAccess(false);
    } else {
      setSubscriptionAccess(data === true);
    }

    setSubscriptionLoading(false);
  }

  function go(name) {
    setPage(name);
    setMobileMenu(false);
  }

  async function logout() {
    await supabase.auth.signOut();
  }

  const pages = {
    dashboard: <Dashboard businessId={businessId} business={business} go={go} refresh={refresh} />,
    customers: <CustomersPage businessId={businessId} refresh={refresh} />,
    analytics: <AnalyticsPage businessId={businessId} refresh={refresh} />,
    cashClosing: <CashClosingPage businessId={businessId} refresh={refresh} />,
    optionalTools: <OptionalToolsPage businessId={businessId} />,
    products: (
      <ProductsPage
        businessId={businessId}
        staff={staff}
        onChanged={() => setRefresh((x) => x + 1)}
      />
    ),
    services: (
      <ServicesPage
        businessId={businessId}
        staff={staff}
        onChanged={() => setRefresh((x) => x + 1)}
      />
    ),
    stock: (
      <StockInPage
        businessId={businessId}
        staff={staff}
        onChanged={() => setRefresh((x) => x + 1)}
      />
    ),
    inventoryHistory: <InventoryHistoryPage businessId={businessId} refresh={refresh} />,
    sales: (
      <SalesPage
        businessId={businessId}
        staff={staff}
        business={business}
        onChanged={() => setRefresh((x) => x + 1)}
      />
    ),
    expenses: (
      <ExpensesPage
        businessId={businessId}
        staff={staff}
        onChanged={() => setRefresh((x) => x + 1)}
      />
    ),
    deposits: (
      <DepositsPage
        businessId={businessId}
        staff={staff}
        onChanged={() => setRefresh((x) => x + 1)}
      />
    ),
    reports: <ReportsPage businessId={businessId} refresh={refresh} />,
    staff: <StaffPage businessId={businessId} refresh={refresh} />,
    credits: <CreditPage businessId={businessId} staff={staff} refresh={refresh} />,
    customerCredits: <CustomerCreditsPage businessId={businessId} staff={staff} refresh={refresh} />,
    payables: <PayablesPage businessId={businessId} staff={staff} refresh={refresh} />,
    attendance: <AttendancePage businessId={businessId} staff={staff} refresh={refresh} />,
    settings: (
      <BusinessSettingsPage
        businessId={businessId}
        business={business}
        user={user}
        theme={theme}
        setTheme={setTheme}
        density={density}
        setDensity={setDensity}
        onChanged={() => { setRefresh((x) => x + 1); loadStaff(); }}
      />
    ),
    saasOverview: <SaaSAdminPage section="overview" refresh={refresh} />,
    saasCustomers: <SaaSAdminPage section="customers" refresh={refresh} />,
    saasSubscriptions: <SaaSAdminPage section="subscriptions" refresh={refresh} />,
    saasPayments: <SaaSAdminPage section="payments" refresh={refresh} />,
  };

  return (
    <div className={`app-shell theme-${theme} density-${density}`} data-theme={theme} data-density={density}>
      {showWelcome && (
        <div className="bless-welcome-backdrop" onClick={closeWelcome}>
          <div className="bless-welcome-card" onClick={(e) => e.stopPropagation()}>
            <div className="bless-welcome-orbit one" />
            <div className="bless-welcome-orbit two" />
            <div className="bless-welcome-logo">B</div>
            <div className="bless-welcome-kicker">BLESS BUSINESS OS</div>
            <h1>Karibu tena, {staff?.staff_name?.split(" ")[0] || "Boss"}.</h1>
            <p className="bless-welcome-sub">{business?.business_name || "Biashara yako"} iko tayari. Mfumo uko hewani.</p>
            <div className="bless-pulse-grid">
              <div><span>SALES</span><strong>{money(pulse.sales)}</strong></div>
              <div><span>EXPENSES</span><strong>{money(pulse.expenses)}</strong></div>
              <div><span>PRODUCTS</span><strong>{pulse.products}</strong></div>
              <div><span>CREDIT</span><strong>{money(pulse.debts)}</strong></div>
            </div>
            <button className="bless-enter-btn" onClick={closeWelcome}>ENTER SYSTEM <span>→</span></button>
            <small>Tip: bonyeza <kbd>Ctrl</kbd> + <kbd>K</kbd> kufungua Quick Command.</small>
          </div>
        </div>
      )}

      {showCommand && (
        <div className="bless-command-backdrop" onClick={() => setShowCommand(false)}>
          <div className="bless-command" onClick={(e) => e.stopPropagation()}>
            <div className="bless-command-head">
              <div><span>QUICK COMMAND</span><h3>Unataka kwenda wapi?</h3></div>
              <button onClick={() => setShowCommand(false)}>×</button>
            </div>
            <div className="bless-command-search">⌘ <span>Chagua action hapa chini</span><kbd>ESC</kbd></div>
            <div className="bless-command-grid">
              {[
                ["sales","🛒","New Sale"],["stock","📦","Stock In"],["inventoryHistory","↕","Inventory History"],
                ["products","▣","Products"],["services","⚙","Services"],["customers","👥","Customers"],
                ["analytics","📊","Profit & Loss"],["cashClosing","💵","Cash Closing"],["reports","▤","Reports"],
                ["credits","💳","Madeni"],["payables","📌","Madeni Yetu"],["customerCredits","💰","Customer Credits"],["attendance","🕐","Attendance"],["settings","⚙","Settings"],
              ].map(([target, icon, label]) => (
                <button key={target} onClick={() => { setShowCommand(false); go(target); }}>
                  <span>{icon}</span><strong>{label}</strong><em>→</em>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      <button className="bless-floating-command" onClick={() => setShowCommand(true)} title="Quick Command">
        <span>⌘</span><small>Ctrl K</small>
      </button>

      <aside className={`sidebar ${mobileMenu ? "open" : ""}`}>
        <div className="sidebar-brand">
          <div className="small-logo">B</div>
          <div>
            <strong>{business?.business_name || "Bless Stationery"}</strong>
            <span>Management System</span>
          </div>
        </div>

        <div className="sidebar-user">
          <div className="avatar">
            {(staff?.staff_name || "A").charAt(0).toUpperCase()}
          </div>
          <div>
            <strong>{staff?.staff_name || "Admin"}</strong>
            <span>{staff?.role || "ADMIN"}</span>
          </div>
        </div>

        <nav>
          <NavButton
            active={page === "dashboard"}
            icon="⌂"
            text="Dashboard"
            onClick={() => go("dashboard")}
          />

          <div className="nav-title">BIASHARA</div>

          <NavButton
            active={page === "sales"}
            icon="🛒"
            text="Sales Entry"
            onClick={() => go("sales")}
          />

          <NavButton
            active={page === "stock"}
            icon="📦"
            text="Stock In"
            onClick={() => go("stock")}
          />

          <NavButton
            active={page === "inventoryHistory"}
            icon="↕"
            text="Inventory History"
            onClick={() => go("inventoryHistory")}
          />

          <NavButton
            active={page === "products"}
            icon="▣"
            text="Products"
            onClick={() => go("products")}
          />

          <NavButton
            active={page === "services"}
            icon="⚙"
            text="Services"
            onClick={() => go("services")}
          />

          <NavButton
            active={page === "customers"}
            icon="👥"
            text="Customers"
            onClick={() => go("customers")}
          />

          <div className="nav-title">FEDHA</div>

          <NavButton
            active={page === "analytics"}
            icon="📊"
            text="Profit & Loss"
            onClick={() => go("analytics")}
          />

          <NavButton
            active={page === "cashClosing"}
            icon="💵"
            text="Daily Cash Closing"
            onClick={() => go("cashClosing")}
          />

          <NavButton
            active={page === "expenses"}
            icon="−"
            text="Expenses"
            onClick={() => go("expenses")}
          />

          <NavButton
            active={page === "deposits"}
            icon="↓"
            text="Deposits"
            onClick={() => go("deposits")}
          />

          <div className="nav-title">TAARIFA</div>

          <NavButton
            active={page === "reports"}
            icon="▤"
            text="Reports"
            onClick={() => go("reports")}
          />

          <NavButton
            active={page === "staff"}
            icon="♟"
            text="Staff"
            onClick={() => go("staff")}
          />

          <NavButton
            active={page === "credits"}
            icon="💳"
            text="Madeni / Credit"
            onClick={() => go("credits")}
          />

          <NavButton
            active={page === "payables"}
            icon="📌"
            text="Madeni Yetu"
            onClick={() => go("payables")}
          />
          <NavButton
            active={page === "customerCredits"}
            icon="💰"
            text="Customer Credits"
            onClick={() => go("customerCredits")}
          />

          <NavButton
            active={page === "attendance"}
            icon="🕐"
            text="Staff Attendance"
            onClick={() => go("attendance")}
          />

          <NavButton
            active={page === "settings"}
            icon="⚙"
            text="Business Settings"
            onClick={() => go("settings")}
          />

          <NavButton
            active={page === "optionalTools"}
            icon="🛠"
            text="Advanced Tools"
            onClick={() => go("optionalTools")}
          />

          {isPlatformAdmin && (
            <>
              <div className="nav-title platform-nav-title">SAAS ADMIN</div>

              <NavButton
                active={page === "saasOverview"}
                icon="◈"
                text="Overview"
                onClick={() => go("saasOverview")}
              />

              <NavButton
                active={page === "saasCustomers"}
                icon="👥"
                text="Customers"
                onClick={() => go("saasCustomers")}
              />

              <NavButton
                active={page === "saasSubscriptions"}
                icon="💳"
                text="Subscriptions"
                onClick={() => go("saasSubscriptions")}
              />

              <NavButton
                active={page === "saasPayments"}
                icon="💰"
                text="Payments"
                onClick={() => go("saasPayments")}
              />
            </>
          )}
        </nav>

        <button className="logout-btn" onClick={logout}>
          ⇥ Logout
        </button>
      </aside>

      {mobileMenu && (
        <div
          className="mobile-overlay"
          onClick={() => setMobileMenu(false)}
        />
      )}

      <main className="main-area">
        <header className="topbar">
          <button
            className="mobile-menu-btn"
            onClick={() => setMobileMenu(!mobileMenu)}
          >
            ☰
          </button>

          <div>
            <strong>
              {page === "dashboard"
                ? "Dashboard"
                : page === "sales"
                ? "Sales Entry"
                : page === "stock"
                ? "Stock In"
                : page === "inventoryHistory"
                ? "Inventory History"
                : page === "products"
                ? "Products"
                : page === "services"
                ? "Services"
                : page === "customers"
                ? "Customers"
                : page === "analytics"
                ? "Profit & Loss"
                : page === "cashClosing"
                ? "Daily Cash Closing"
                : page === "optionalTools"
                ? "Advanced Tools"
                : page === "expenses"
                ? "Expenses"
                : page === "deposits"
                ? "Deposits"
                : page === "reports"
                ? "Reports"
                : page === "staff"
                ? "Staff"
                : page === "credits"
                ? "Madeni / Credit"
                : page === "payables"
                ? "Madeni Yetu"
                : page === "customerCredits"
                ? "Customer Credits"
                : page === "attendance"
                ? "Staff Attendance"
                : page === "settings"
                ? "Business Settings"
                : page === "saasOverview"
                ? "SaaS Overview"
                : page === "saasCustomers"
                ? "Customers"
                : page === "saasSubscriptions"
                ? "Subscriptions"
                : page === "saasPayments"
                ? "Payments"
                : "Staff"}
            </strong>
            <span className="topbar-date">
              {new Date().toLocaleDateString("en-TZ", {
                weekday: "long",
                year: "numeric",
                month: "long",
                day: "numeric",
              })}
            </span>
          </div>

          <div className="top-actions">
            <div className="language-control" title="Lugha">
              <span className="language-icon">🌐</span>
              <span className="language-label">Lugha</span>
              <select
                className="language-select"
                value={language}
                onChange={(e) => changeLanguage(e.target.value)}
                aria-label="Lugha"
              >
                {(availableLanguages.length ? availableLanguages : [{ value: "sw", label: "Kiswahili" }, { value: "en", label: "English" }, { value: "nl", label: "Dutch" }, { value: "fr", label: "French" }, { value: "de", label: "German" }, { value: "es", label: "Spanish" }]).map((item) => (
                  <option key={item.value} value={item.value}>{item.label}</option>
                ))}
              </select>
              <div id="google_translate_element" className="google-translate-hidden" />
            </div>
            <button className="refresh-btn" onClick={() => setRefresh((x) => x + 1)}>
              ↻ Refresh
            </button>
            <div className="user-pill">
              <span className="online-dot" />
              {staff?.staff_name || user.email}
            </div>
          </div>
        </header>

        <section className="content">
          {isPlatformAdmin && page.startsWith("saas") ? (
            pages[page]
          ) : !businessId ? (
            <div className="panel">
              <h3>Business profile haijapatikana</h3>
              <p>Account hii haijaunganishwa na biashara. Wasiliana na administrator.</p>
            </div>
          ) : subscriptionLoading ? (
            <div className="page-loading">
              <div className="spinner" />
              <p>Inathibitisha subscription...</p>
            </div>
          ) : !isPlatformAdmin && subscriptionAccess === false ? (
            <SubscriptionRequiredPage
              business={business}
              onRefresh={() => checkSubscriptionAccess(businessId)}
            />
          ) : (
            pages[page]
          )}
        </section>
      </main>
    </div>
  );
}

function NavButton({ active, icon, text, onClick }) {
  return (
    <button className={`nav-button ${active ? "active" : ""}`} onClick={onClick}>
      <span className="nav-icon">{icon}</span>
      <span>{text}</span>
    </button>
  );
}


function SubscriptionRequiredPage({ business, onRefresh }) {
  const [payingPlan, setPayingPlan] = useState("");
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");

  async function startPayment(plan) {
    setError("");
    setMessage("");
    setPayingPlan(plan);

    try {
      const { data, error: functionError } = await supabase.functions.invoke(
        "super-function",
        { body: { action: "create-payment", plan } }
      );

      if (functionError) {
        throw new Error(functionError.message || "Payment request imeshindikana.");
      }

      const checkoutUrl = data?.checkout_url || data?.redirect_url;
      if (!checkoutUrl) {
        throw new Error(data?.message || "Pesapal haikurudisha checkout URL.");
      }

      setMessage("Inafungua ukurasa wa malipo wa Pesapal...");
      window.location.href = checkoutUrl;
    } catch (err) {
      console.error("Pesapal payment start failed:", err);
      setError(err?.message || "Imeshindikana kuanzisha malipo. Jaribu tena.");
      setPayingPlan("");
    }
  }

  return (
    <div className="panel subscription-gate">
      <div className="subscription-gate-icon">💳</div>
      <h2>Subscription inahitajika</h2>
      <p>
        Muda wa FREE TRIAL wa <strong>{business?.business_name || "biashara yako"}</strong>
        umeisha au subscription haijawa ACTIVE.
      </p>
      <p className="muted">
        Chagua mpango wako na ukamilishe malipo kupitia Pesapal.
      </p>

      {message && <div className="success-box">{message}</div>}
      {error && <div className="error-box">{error}</div>}

      <div className="subscription-gate-actions" style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        <button
          className="primary-btn"
          disabled={!!payingPlan}
          onClick={() => startPayment("MONTHLY")}
        >
          {payingPlan === "MONTHLY" ? "Inaandaa malipo..." : "LIPA MWEZI — TZS 10,000"}
        </button>
        <button
          className="secondary-btn"
          disabled={!!payingPlan}
          onClick={() => startPayment("YEARLY")}
        >
          {payingPlan === "YEARLY" ? "Inaandaa malipo..." : "LIPA MWAKA — TZS 50,000"}
        </button>
        <button
          className="secondary-btn"
          disabled={!!payingPlan}
          onClick={onRefresh}
        >
          ↻ Angalia Tena
        </button>
      </div>
    </div>
  );
}

function SaaSAdminPage({ section = "overview", refresh }) {
  const [businesses, setBusinesses] = useState([]);
  const [staffRows, setStaffRows] = useState([]);
  const [subscriptions, setSubscriptions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("ALL");
  const [planFilter, setPlanFilter] = useState("ALL");
  const [selectedBusinessId, setSelectedBusinessId] = useState(null);

  useEffect(() => {
    load();
  }, [refresh]);

  async function load() {
    setLoading(true);

    const [businessResult, staffResult, subscriptionResult] = await Promise.all([
      supabase
        .from("businesses")
        .select("*")
        .order("created_at", { ascending: false }),
      supabase
        .from("staff")
        .select("id, business_id, staff_name, role, user_id, active")
        .eq("role", "ADMIN"),
      supabase
        .from("subscriptions")
        .select("*")
        .order("created_at", { ascending: false }),
    ]);

    if (businessResult.error) {
      console.error("SaaS businesses error:", businessResult.error);
      alert(businessResult.error.message);
    }

    if (staffResult.error) {
      console.error("SaaS staff error:", staffResult.error);
    }

    if (subscriptionResult.error) {
      console.error("SaaS subscriptions error:", subscriptionResult.error);
    }

    setBusinesses(businessResult.data || []);
    setStaffRows(staffResult.data || []);
    setSubscriptions(subscriptionResult.data || []);
    setLoading(false);
  }

  const ownerMap = useMemo(() => {
    const map = {};
    staffRows.forEach((row) => {
      if (!map[row.business_id]) map[row.business_id] = row.staff_name;
    });
    return map;
  }, [staffRows]);

  const latestSubscriptionMap = useMemo(() => {
    const map = {};
    subscriptions.forEach((row) => {
      if (!map[row.business_id]) map[row.business_id] = row;
    });
    return map;
  }, [subscriptions]);

  const customerRows = useMemo(() => {
    const q = search.trim().toLowerCase();

    return businesses.filter((business) => {
      const latest = latestSubscriptionMap[business.id];
      const status = latest?.status || business.subscription_status || "-";
      const plan = latest?.plan || business.plan || "-";
      const owner = ownerMap[business.id] || "-";

      const matchesSearch = !q || [
        business.business_name,
        owner,
        business.email,
        business.phone,
        business.country,
        business.region,
        business.district,
        business.ward,
        business.street,
      ].some((value) => String(value || "").toLowerCase().includes(q));

      const matchesStatus = statusFilter === "ALL" || status === statusFilter;
      const matchesPlan = planFilter === "ALL" || plan === planFilter;

      return matchesSearch && matchesStatus && matchesPlan;
    });
  }, [businesses, ownerMap, latestSubscriptionMap, search, statusFilter, planFilter]);

  const stats = useMemo(() => {
    const rows = businesses.map((business) => {
      const latest = latestSubscriptionMap[business.id];
      return latest?.status || business.subscription_status || "-";
    });

    return {
      total: businesses.length,
      active: rows.filter((x) => x === "ACTIVE").length,
      trialing: rows.filter((x) => x === "TRIALING").length,
      pending: rows.filter((x) => x === "PENDING_PAYMENT").length,
      expired: rows.filter((x) => x === "EXPIRED").length,
    };
  }, [businesses, latestSubscriptionMap]);

  const selectedBusiness = businesses.find((x) => x.id === selectedBusinessId) || null;
  const selectedSubscriptions = selectedBusinessId
    ? subscriptions.filter((x) => x.business_id === selectedBusinessId)
    : [];

  function statusLabel(status) {
    const labels = {
      ACTIVE: "ACTIVE",
      TRIALING: "TRIALING",
      PENDING_PAYMENT: "PENDING PAYMENT",
      EXPIRED: "EXPIRED",
      CANCELLED: "CANCELLED",
    };
    return labels[status] || status || "-";
  }

  function planLabel(plan) {
    const labels = {
      FREE_TRIAL: "FREE TRIAL",
      MONTHLY: "MWEZI",
      YEARLY: "MWAKA",
    };
    return labels[plan] || plan || "-";
  }

  function statusClass(status) {
    return `saas-status ${String(status || "unknown").toLowerCase().replace(/_/g, "-")}`;
  }

  if (loading) return <PageLoading />;

  if (section === "overview") {
    return (
      <div>
        <PageTitle
          title="SaaS Overview"
          subtitle="Muhtasari wa businesses zote zilizo kwenye platform yako."
        />

        <div className="stats-grid">
          <StatCard title="Customers Wote" value={stats.total} icon="👥" tone="blue" />
          <StatCard title="Active" value={stats.active} icon="✓" tone="green" />
          <StatCard title="Free Trial" value={stats.trialing} icon="⏱" tone="purple" />
          <StatCard title="Pending Payment" value={stats.pending} icon="💳" tone="orange" />
        </div>

        <div className="stats-grid small-stats">
          <StatCard title="Expired" value={stats.expired} icon="⚠" tone="red" />
          <StatCard title="Mwezi" value={businesses.filter((b) => (latestSubscriptionMap[b.id]?.plan || b.plan) === "MONTHLY").length} icon="📅" />
          <StatCard title="Mwaka" value={businesses.filter((b) => (latestSubscriptionMap[b.id]?.plan || b.plan) === "YEARLY").length} icon="🏆" />
          <StatCard title="Subscriptions" value={subscriptions.length} icon="▣" />
        </div>

        <div className="panel">
          <div className="panel-header">
            <div>
              <h3>Registrations za Karibuni</h3>
              <span>Businesses mpya zilizojiunga na platform</span>
            </div>
            <button className="text-btn" onClick={() => setSelectedBusinessId(null)}>
              {stats.total} Customers
            </button>
          </div>

          <div className="saas-table-wrap">
            <table className="saas-table">
              <thead>
                <tr>
                  <th>Business</th>
                  <th>Owner</th>
                  <th>Plan</th>
                  <th>Status</th>
                  <th>Registered</th>
                </tr>
              </thead>
              <tbody>
                {businesses.slice(0, 10).map((business) => {
                  const latest = latestSubscriptionMap[business.id];
                  return (
                    <tr key={business.id} onClick={() => setSelectedBusinessId(business.id)}>
                      <td><strong>{business.business_name}</strong></td>
                      <td>{ownerMap[business.id] || "-"}</td>
                      <td>{planLabel(latest?.plan || business.plan)}</td>
                      <td><span className={statusClass(latest?.status || business.subscription_status)}>{statusLabel(latest?.status || business.subscription_status)}</span></td>
                      <td>{formatDate(business.created_at)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {!businesses.length && <EmptyState text="Bado hakuna business iliyojisajili." />}
        </div>
      </div>
    );
  }

  if (section === "customers") {
    return (
      <div>
        <PageTitle
          title="Customers"
          subtitle="Orodha ya businesses zote zilizojiunga na Bless Business SaaS."
        />

        <div className="saas-toolbar">
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search business, owner, email, phone, location..."
          />
          <select value={planFilter} onChange={(e) => setPlanFilter(e.target.value)}>
            <option value="ALL">Plans zote</option>
            <option value="FREE_TRIAL">FREE TRIAL</option>
            <option value="MONTHLY">MWEZI</option>
            <option value="YEARLY">MWAKA</option>
          </select>
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
            <option value="ALL">Status zote</option>
            <option value="TRIALING">TRIALING</option>
            <option value="PENDING_PAYMENT">PENDING PAYMENT</option>
            <option value="ACTIVE">ACTIVE</option>
            <option value="EXPIRED">EXPIRED</option>
            <option value="CANCELLED">CANCELLED</option>
          </select>
          <button className="secondary-btn" onClick={load}>↻ Refresh</button>
        </div>

        <div className="panel">
          <div className="panel-header">
            <div>
              <h3>{customerRows.length} Customer{customerRows.length === 1 ? "" : "s"}</h3>
              <span>Bonyeza customer kuona maelezo yake.</span>
            </div>
          </div>

          <div className="saas-table-wrap">
            <table className="saas-table">
              <thead>
                <tr>
                  <th>Business</th>
                  <th>Owner</th>
                  <th>Email</th>
                  <th>Phone</th>
                  <th>Location</th>
                  <th>Plan</th>
                  <th>Status</th>
                  <th>Registered</th>
                </tr>
              </thead>
              <tbody>
                {customerRows.map((business) => {
                  const latest = latestSubscriptionMap[business.id];
                  return (
                    <tr key={business.id} onClick={() => setSelectedBusinessId(business.id)}>
                      <td><strong>{business.business_name}</strong></td>
                      <td>{ownerMap[business.id] || "-"}</td>
                      <td>{business.email || "-"}</td>
                      <td>{business.phone || "-"}</td>
                      <td>
                        {[business.country, business.region, business.district].filter(Boolean).join(", ") || "-"}
                      </td>
                      <td>{planLabel(latest?.plan || business.plan)}</td>
                      <td><span className={statusClass(latest?.status || business.subscription_status)}>{statusLabel(latest?.status || business.subscription_status)}</span></td>
                      <td>{formatDate(business.created_at)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {!customerRows.length && <EmptyState text="Hakuna customer anayefanana na search/filter yako." />}
        </div>

        {selectedBusiness && (
          <div className="panel saas-detail-panel">
            <div className="panel-header">
              <div>
                <h3>{selectedBusiness.business_name}</h3>
                <span>Customer details</span>
              </div>
              <button className="secondary-btn" onClick={() => setSelectedBusinessId(null)}>Funga</button>
            </div>

            <div className="saas-detail-grid">
              <div><span>Owner</span><strong>{ownerMap[selectedBusiness.id] || "-"}</strong></div>
              <div><span>Email</span><strong>{selectedBusiness.email || "-"}</strong></div>
              <div><span>Phone</span><strong>{selectedBusiness.phone || "-"}</strong></div>
              <div><span>Country</span><strong>{selectedBusiness.country || "-"}</strong></div>
              <div><span>Region</span><strong>{selectedBusiness.region || "-"}</strong></div>
              <div><span>District / City</span><strong>{selectedBusiness.district || "-"}</strong></div>
              <div><span>Ward</span><strong>{selectedBusiness.ward || "-"}</strong></div>
              <div><span>Street</span><strong>{selectedBusiness.street || "-"}</strong></div>
              <div><span>Registered</span><strong>{formatDate(selectedBusiness.created_at)}</strong></div>
              <div><span>Plan</span><strong>{planLabel(selectedBusiness.plan)}</strong></div>
              <div><span>Status</span><strong>{statusLabel(selectedBusiness.subscription_status)}</strong></div>
            </div>

            <h4 className="saas-subtitle">Subscription History</h4>
            <div className="saas-table-wrap">
              <table className="saas-table">
                <thead>
                  <tr><th>Plan</th><th>Amount</th><th>Status</th><th>Provider</th><th>Reference</th><th>Date</th></tr>
                </thead>
                <tbody>
                  {selectedSubscriptions.map((row) => (
                    <tr key={row.id}>
                      <td>{planLabel(row.plan)}</td>
                      <td>{money(row.amount)} {row.currency || "TZS"}</td>
                      <td><span className={statusClass(row.status)}>{statusLabel(row.status)}</span></td>
                      <td>{row.payment_provider || "-"}</td>
                      <td>{row.payment_reference || row.provider_transaction_id || "-"}</td>
                      <td>{formatDate(row.created_at)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    );
  }

  const paymentRows = subscriptions.filter((row) => row.payment_reference || row.provider_transaction_id || row.payment_provider || number(row.amount) > 0);

  if (section === "subscriptions") {
    return (
      <div>
        <PageTitle title="Subscriptions" subtitle="Subscription status na history ya businesses zote." />
        <div className="panel">
          <div className="saas-table-wrap">
            <table className="saas-table">
              <thead><tr><th>Business</th><th>Plan</th><th>Amount</th><th>Status</th><th>Start</th><th>End / Trial End</th><th>Created</th></tr></thead>
              <tbody>
                {subscriptions.map((row) => {
                  const business = businesses.find((b) => b.id === row.business_id);
                  return (
                    <tr key={row.id}>
                      <td><strong>{business?.business_name || "-"}</strong></td>
                      <td>{planLabel(row.plan)}</td>
                      <td>{money(row.amount)} {row.currency || "TZS"}</td>
                      <td><span className={statusClass(row.status)}>{statusLabel(row.status)}</span></td>
                      <td>{formatDate(row.started_at)}</td>
                      <td>{formatDate(row.ends_at || row.trial_ends_at)}</td>
                      <td>{formatDate(row.created_at)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {!subscriptions.length && <EmptyState text="Hakuna subscription records." />}
        </div>
      </div>
    );
  }

  return (
    <div>
      <PageTitle title="Payments" subtitle="Payment records zilizopo kwa sasa. Flutterwave tutaunganisha baadaye." />
      <div className="panel">
        <div className="saas-table-wrap">
          <table className="saas-table">
            <thead><tr><th>Business</th><th>Plan</th><th>Amount</th><th>Provider</th><th>Reference</th><th>Status</th><th>Date</th></tr></thead>
            <tbody>
              {paymentRows.map((row) => {
                const business = businesses.find((b) => b.id === row.business_id);
                return (
                  <tr key={row.id}>
                    <td><strong>{business?.business_name || "-"}</strong></td>
                    <td>{planLabel(row.plan)}</td>
                    <td>{money(row.amount)} {row.currency || "TZS"}</td>
                    <td>{row.payment_provider || "-"}</td>
                    <td>{row.payment_reference || row.provider_transaction_id || "-"}</td>
                    <td><span className={statusClass(row.status)}>{statusLabel(row.status)}</span></td>
                    <td>{formatDate(row.created_at)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        {!paymentRows.length && <EmptyState text="Bado hakuna payment records. Hii itajaa tukishaunganisha payment gateway." />}
      </div>
    </div>
  );
}

function Dashboard({ businessId, business, go, refresh }) {
  const [stats, setStats] = useState({
    todaySales: 0,
    todayProfit: 0,
    monthSales: 0,
    monthProfit: 0,
    products: 0,
    services: 0,
    lowStock: 0,
    stockValue: 0,
  });
  const [lowProducts, setLowProducts] = useState([]);
  const [allLowProducts, setAllLowProducts] = useState([]);
  const [recentSales, setRecentSales] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    load();
  }, [refresh, businessId]);

  async function load() {
    if (!businessId) return;
    setLoading(true);

    const [
      { data: products },
      { data: services },
      { data: sales },
      { data: stockIn },
    ] = await Promise.all([
      supabase.from("products").select("*").eq("business_id", businessId).eq("active", true),
      supabase.from("services").select("*").eq("business_id", businessId).eq("active", true),
      supabase.from("sales").select("*").eq("business_id", businessId).order("sale_date", { ascending: false }).limit(1000),
      supabase.from("stock_in").select("*").eq("business_id", businessId),
    ]);

    const p = products || [];
    const s = sales || [];
    const si = stockIn || [];

    const stockMap = {};
    p.forEach((x) => {
      stockMap[x.id] = number(x.opening_qty);
    });

    si.forEach((x) => {
      if (stockMap[x.product_id] !== undefined) {
        stockMap[x.product_id] += number(x.quantity_in);
      }
    });

    s.forEach((x) => {
      if (x.product_id && stockMap[x.product_id] !== undefined) {
        stockMap[x.product_id] -= number(x.quantity);
      }
    });

    const low = p
      .map((x) => ({
        ...x,
        currentStock: stockMap[x.id] || 0,
      }))
      .filter((x) => x.currentStock <= number(x.reorder_level))
      .sort((a, b) => a.currentStock - b.currentStock);

    const now = new Date();
    const startToday = new Date();
    startToday.setHours(0, 0, 0, 0);

    const startMonth = new Date();
    startMonth.setDate(1);
    startMonth.setHours(0, 0, 0, 0);

    const todaySales = s.filter((x) => new Date(x.sale_date) >= startToday);
    const monthSales = s.filter((x) => new Date(x.sale_date) >= startMonth);

    setStats({
      todaySales: todaySales.reduce((a, x) => a + number(x.sales_total), 0),
      todayProfit: todaySales.reduce((a, x) => a + number(x.profit), 0),
      monthSales: monthSales.reduce((a, x) => a + number(x.sales_total), 0),
      monthProfit: monthSales.reduce((a, x) => a + number(x.profit), 0),
      products: p.length,
      services: (services || []).length,
      lowStock: low.length,
      stockValue: p.reduce(
        (a, x) =>
          a +
          Math.max(stockMap[x.id] || 0, 0) * number(x.cost_per_each),
        0
      ),
    });

    setAllLowProducts(low);
    setLowProducts(low.slice(0, 8));
    setRecentSales(s.slice(0, 8));
    setLoading(false);
  }

  function printLowStockPurchaseList() {
    if (!allLowProducts.length) {
      alert("Hakuna bidhaa yenye low stock kwa sasa.");
      return;
    }

    const rows = allLowProducts.map((x) => {
      const current = Math.max(number(x.currentStock), 0);
      const reorder = Math.max(number(x.reorder_level), 0);
      const qtyToBuy = Math.max(reorder - current, 0);
      const unitCost = Math.max(number(x.cost_per_each), 0);
      return {
        name: x.product_name,
        unit: x.unit || "PCS",
        current,
        reorder,
        qtyToBuy,
        unitCost,
        total: qtyToBuy * unitCost,
      };
    });

    const totalCost = rows.reduce((sum, row) => sum + row.total, 0);
    const printWindow = window.open("", "_blank", "width=1000,height=800");

    if (!printWindow) {
      alert("Browser imezuia print window. Ruhusu pop-ups kisha ujaribu tena.");
      return;
    }

    const generatedAt = new Date().toLocaleString("en-TZ", {
      dateStyle: "medium",
      timeStyle: "short",
    });
    const businessName = business?.business_name || "Bless Stationery";

    printWindow.document.write(`<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Bless Stationery - Low Stock Purchase List</title>
  <style>
    *{box-sizing:border-box}
    body{font-family:Arial,sans-serif;color:#111;padding:28px;font-size:12px}
    h1{margin:0 0 6px;font-size:22px}
    h2{margin:0 0 4px;font-size:16px}
    .muted{color:#666;margin-bottom:18px}
    table{width:100%;border-collapse:collapse;margin-top:18px}
    th,td{border:1px solid #ccc;padding:8px;text-align:left}
    th{background:#f2f4f7}
    .num{text-align:right}
    tfoot td{font-weight:700;font-size:14px;background:#fafafa}
    .note{margin-top:16px;padding:10px;border:1px solid #ddd;background:#fafafa}
    @media print{body{padding:0}.no-print{display:none}}
  </style>
</head>
<body>
  <h1>${escapeHtml(businessName)}</h1>
  <h2>Low Stock Purchase List</h2>
  <div class="muted">Generated: ${generatedAt}</div>
  <table>
    <thead>
      <tr>
        <th>#</th>
        <th>Product</th>
        <th>Unit</th>
        <th class="num">Current Stock</th>
        <th class="num">Reorder Level</th>
        <th class="num">Qty to Buy</th>
        <th class="num">Cost / Each</th>
        <th class="num">Estimated Cost</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map((row, i) => `
        <tr>
          <td>${i + 1}</td>
          <td>${String(row.name || "").replace(/[&<>\"]/g, (m) => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[m]))}</td>
          <td>${row.unit}</td>
          <td class="num">${row.current}</td>
          <td class="num">${row.reorder}</td>
          <td class="num"><strong>${row.qtyToBuy}</strong></td>
          <td class="num">${money(row.unitCost)}</td>
          <td class="num">${money(row.total)}</td>
        </tr>
      `).join("")}
    </tbody>
    <tfoot>
      <tr>
        <td colspan="7" class="num">TOTAL ESTIMATED PURCHASE COST</td>
        <td class="num">${money(totalCost)}</td>
      </tr>
    </tfoot>
  </table>
  <div class="note">Qty to Buy imehesabiwa kufikisha kila bidhaa kwenye Reorder Level. Rekebisha kiasi kabla ya kununua ikiwa mahitaji ya biashara yanahitaji zaidi.</div>
  <script>window.onload=function(){window.print();}</script>
</body>
</html>`);

    printWindow.document.close();
    setTimeout(() => {
      try {
        printWindow.focus();
        printWindow.print();
      } catch (e) {}
    }, 300);
  }

  if (loading) return <PageLoading />;

  return (
    <div>
      <div className="welcome">
        <div>
          <h1>Karibu kwenye Bless Stationery 👋</h1>
          <p>Hapa utaona hali ya biashara yako kwa haraka.</p>
        </div>

        <div className="quick-actions">
          <button className="primary-btn" onClick={() => go("sales")}>
            + New Sale
          </button>
          <button className="secondary-btn" onClick={() => go("stock")}>
            + Stock In
          </button>
        </div>
      </div>

      <div className="stats-grid">
        <StatCard
          title="Sales za Leo"
          value={money(stats.todaySales)}
          icon="💰"
          tone="blue"
        />
        <StatCard
          title="Profit ya Leo"
          value={money(stats.todayProfit)}
          icon="📈"
          tone="green"
        />
        <StatCard
          title="Sales za Mwezi"
          value={money(stats.monthSales)}
          icon="🧾"
          tone="purple"
        />
        <StatCard
          title="Profit ya Mwezi"
          value={money(stats.monthProfit)}
          icon="💎"
          tone="orange"
        />
      </div>

      <div className="stats-grid small-stats">
        <StatCard title="Products" value={stats.products} icon="📦" />
        <StatCard title="Services" value={stats.services} icon="⚙" />
        <StatCard
          title="Low Stock"
          value={stats.lowStock}
          icon="⚠"
          tone={stats.lowStock ? "red" : "green"}
        />
        <StatCard
          title="Stock Value"
          value={money(stats.stockValue)}
          icon="🏷"
        />
      </div>

      <div className="dashboard-grid">
        <div className="panel">
          <div className="panel-header">
            <div>
              <h3>Low Stock Alert</h3>
              <span>Bidhaa zinazohitaji kuongezewa</span>
            </div>
            <div className="panel-header-actions">
              <button className="secondary-btn small-btn" onClick={printLowStockPurchaseList}>
                🖨 Print Purchase List
              </button>
              <button className="text-btn" onClick={() => go("products")}>
                View all
              </button>
            </div>
          </div>

          {lowProducts.length === 0 ? (
            <EmptyState text="Hakuna bidhaa yenye low stock." />
          ) : (
            <div className="mini-table">
              {lowProducts.map((x) => (
                <div className="mini-row" key={x.id}>
                  <div>
                    <strong>{x.product_name}</strong>
                    <span>{x.unit}</span>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div
                      className={
                        x.currentStock <= 0
                          ? "stock-danger"
                          : "stock-warning"
                      }
                    >
                      {x.currentStock} {x.unit}
                    </div>
                    <div style={{ fontSize: 12, color: "#666", marginTop: 3 }}>
                      Reorder: {x.reorder_level} {x.unit}
                    </div>
                    <div style={{ fontSize: 12, color: "#b42318", fontWeight: 700, marginTop: 2 }}>
                      Ongeza: {Math.max(number(x.reorder_level) - number(x.currentStock), 0)} {x.unit}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="panel">
          <div className="panel-header">
            <div>
              <h3>Sales za Karibuni</h3>
              <span>Transactions za mwisho</span>
            </div>
            <button className="text-btn" onClick={() => go("sales")}>
              View all
            </button>
          </div>

          {recentSales.length === 0 ? (
            <EmptyState text="Bado hakuna sales." />
          ) : (
            <div className="mini-table">
              {recentSales.map((x) => (
                <div className="mini-row" key={x.id}>
                  <div>
                    <strong>
                      {x.sale_type === "SERVICE"
                        ? "Service"
                        : x.product_id
                        ? "Product"
                        : "Manual Item"}
                    </strong>
                    <span>{formatDate(x.sale_date)}</span>
                  </div>
                  <strong>{money(x.sales_total)}</strong>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function StatCard({ title, value, icon, tone = "blue" }) {
  return (
    <div className={`stat-card ${tone}`}>
      <div>
        <span className="stat-title">{title}</span>
        <strong>{value}</strong>
      </div>
      <div className="stat-icon">{icon}</div>
    </div>
  );
}

function ProductsPage({ businessId, onChanged }) {
  const [products, setProducts] = useState([]);
  const [stockIn, setStockIn] = useState([]);
  const [sales, setSales] = useState([]);
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [stockFilter, setStockFilter] = useState("ALL");
  const [busy, setBusy] = useState(false);

  const [form, setForm] = useState({
    product_name: "",
    unit: "PCS",
    opening_qty: "0",
    cost_per_each: "",
    sell_per_each: "",
    reorder_level: "5",
  });

  useEffect(() => {
    load();
  }, [businessId]);

  useEffect(() => {
    let cancelled = false;
    async function loadCustomerCredit() {
      if (!businessId || form.payment_method !== "CUSTOMER_CREDIT" || !form.customer_name.trim()) {
        setAvailableCustomerCredit(0);
        setCustomerCreditRows([]);
        return;
      }
      const { data, error } = await supabase
        .from("customer_credits")
        .select("*")
        .eq("business_id", businessId)
        .ilike("customer_name", form.customer_name.trim())
        .gt("balance", 0)
        .order("created_at", { ascending: true });
      if (cancelled) return;
      if (error) {
        console.error("Failed to load customer credit:", error);
        setAvailableCustomerCredit(0);
        setCustomerCreditRows([]);
        return;
      }
      const rows = data || [];
      setCustomerCreditRows(rows);
      setAvailableCustomerCredit(rows.reduce((a, x) => a + number(x.balance), 0));
    }
    loadCustomerCredit();
    return () => { cancelled = true; };
  }, [businessId, form.payment_method, form.customer_name, form.customer_phone]);

  async function load() {
    if (!businessId) return;

    const [{ data: p }, { data: si }, { data: s }] = await Promise.all([
      supabase.from("products").select("*").eq("business_id", businessId).order("product_name"),
      supabase.from("stock_in").select("*").eq("business_id", businessId),
      supabase.from("sales").select("product_id, quantity").eq("business_id", businessId),
    ]);

    setProducts(p || []);
    setStockIn(si || []);
    setSales(s || []);
  }

  const rows = useMemo(() => {
    return products
      .map((p) => {
        const incoming = stockIn
          .filter((x) => x.product_id === p.id)
          .reduce((a, x) => a + number(x.quantity_in), 0);

        const sold = sales
          .filter((x) => x.product_id === p.id)
          .reduce((a, x) => a + number(x.quantity), 0);

        return {
          ...p,
          currentStock: number(p.opening_qty) + incoming - sold,
        };
      })
      .filter((x) =>
        x.product_name.toLowerCase().includes(search.toLowerCase())
      )
      .filter((x) => {
        if (stockFilter === "LOW") return x.currentStock <= number(x.reorder_level);
        if (stockFilter === "OUT") return x.currentStock <= 0;
        if (stockFilter === "OK") return x.currentStock > number(x.reorder_level);
        return true;
      });
  }, [products, stockIn, sales, search]);

  function resetProductForm() {
    setForm({
      product_name: "",
      unit: "PCS",
      opening_qty: "0",
      cost_per_each: "",
      sell_per_each: "",
      reorder_level: "5",
    });
    setEditingId(null);
  }

  function startEditProduct(product) {
    setEditingId(product.id);
    setForm({
      product_name: product.product_name || "",
      unit: product.unit || "PCS",
      opening_qty: String(product.opening_qty ?? 0),
      cost_per_each: String(product.cost_per_each ?? ""),
      sell_per_each: String(product.sell_per_each ?? ""),
      reorder_level: String(product.reorder_level ?? 5),
    });
    setShowForm(true);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function addProduct(e) {
    e.preventDefault();

    if (!form.product_name.trim()) {
      alert("Weka jina la bidhaa.");
      return;
    }
    if (positiveNumber(form.cost_per_each) < 0 || positiveNumber(form.sell_per_each) < 0) {
      alert("Bei haiwezi kuwa chini ya 0.");
      return;
    }

    setBusy(true);

    const payload = {
      product_name: form.product_name.trim(),
      unit: form.unit.trim() || "PCS",
      opening_qty: positiveNumber(form.opening_qty),
      cost_per_each: positiveNumber(form.cost_per_each),
      sell_per_each: positiveNumber(form.sell_per_each),
      reorder_level: positiveNumber(form.reorder_level),
      active: true,
    };

    const result = editingId
      ? await supabase.from("products").update(payload).eq("id", editingId).eq("business_id", businessId)
      : await supabase.from("products").insert({ business_id: businessId, ...payload });

    if (result.error) {
      alert(result.error.message.includes("duplicate") ? "Bidhaa hiyo tayari ipo." : result.error.message);
    } else {
      resetProductForm();
      setShowForm(false);
      await load();
      onChanged();
    }

    setBusy(false);
  }

  async function deleteProduct(id, name) {
    if (!confirm(`Una uhakika unataka kufuta "${name}"?`)) return;

    const { error } = await supabase
      .from("products")
      .delete()
      .eq("id", id)
      .eq("business_id", businessId);

    if (error) {
      alert(
        "Bidhaa hii haiwezi kufutwa kwa sababu tayari ina Stock In au Sales. Tumia records zake kwanza."
      );
    } else {
      await load();
      onChanged();
    }
  }

  return (
    <div>
      <PageTitle
        title="Products"
        subtitle="Simamia bidhaa, bei na stock."
        button={showForm ? "Close Form" : "+ Add Product"}
        onClick={() => { if (showForm) { setShowForm(false); resetProductForm(); } else setShowForm(true); }}
      />

      {showForm && (
        <div className="panel form-panel">
          <div className="panel-header">
            <div>
              <h3>{editingId ? "Hariri Bidhaa" : "Ongeza Bidhaa Mpya"}</h3>
              <span>{editingId ? "Sasisha taarifa za bidhaa bila kupoteza stock history." : "Jaza taarifa za bidhaa"}</span>
            </div>
          </div>

          <form onSubmit={addProduct} className="form-grid">
            <Field
              label="Product Name"
              value={form.product_name}
              onChange={(v) => setForm({ ...form, product_name: v })}
              required
            />
            <Field
              label="Unit"
              value={form.unit}
              onChange={(v) => setForm({ ...form, unit: v })}
            />
            <Field
              label="Opening Qty"
              type="number"
              value={form.opening_qty}
              onChange={(v) => setForm({ ...form, opening_qty: v })}
            />
            <Field
              label="Cost Per Each"
              type="number"
              value={form.cost_per_each}
              onChange={(v) => setForm({ ...form, cost_per_each: v })}
              required
            />
            <Field
              label="Selling Price"
              type="number"
              value={form.sell_per_each}
              onChange={(v) => setForm({ ...form, sell_per_each: v })}
              required
            />
            <Field
              label="Reorder Level"
              type="number"
              value={form.reorder_level}
              onChange={(v) => setForm({ ...form, reorder_level: v })}
            />

            <div className="form-actions full">
              <button className="primary-btn" disabled={busy}>
                {busy ? "Inahifadhi..." : editingId ? "Update Product" : "Save Product"}
              </button>
              <button
                type="button"
                className="secondary-btn"
                onClick={() => setShowForm(false)}
              >
                Cancel
              </button>
            </div>
          </form>
        </div>
      )}

      <div className="panel">
        <div className="toolbar">
          <div className="search-box">
            🔎
            <input
              placeholder="Search product..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
          <div className="toolbar-actions">
            <select value={stockFilter} onChange={(e) => setStockFilter(e.target.value)}>
              <option value="ALL">All Stock</option>
              <option value="LOW">Low Stock</option>
              <option value="OUT">Out of Stock</option>
              <option value="OK">Stock OK</option>
            </select>
            <button className="secondary-btn small-btn" onClick={() => {
              const exportRows = rows.map((p) => ({
                Product: p.product_name,
                Unit: p.unit,
                "Current Stock": p.currentStock,
                "Reorder Level": p.reorder_level,
                "Cost / Each": p.cost_per_each,
                "Selling Price": p.sell_per_each,
              }));
              exportToExcel(exportRows, `Bless-Stationery-Products-${todayDateInput()}.xlsx`, "Products");
            }}>⬇ Export</button>
            <span className="record-count">{rows.length} products</span>
          </div>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Product</th>
                <th>Unit</th>
                <th>Stock</th>
                <th>Cost</th>
                <th>Selling</th>
                <th>Reorder</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((p) => (
                <tr key={p.id}>
                  <td>
                    <strong>{p.product_name}</strong>
                  </td>
                  <td>{p.unit}</td>
                  <td>
                    <StockBadge
                      stock={p.currentStock}
                      reorder={p.reorder_level}
                    />
                  </td>
                  <td>{money(p.cost_per_each)}</td>
                  <td>{money(p.sell_per_each)}</td>
                  <td>{p.reorder_level}</td>
                  <td>
                    <div className="row-actions">
                      <button className="secondary-btn small-btn" onClick={() => startEditProduct(p)}>Edit</button>
                      <button
                        className="danger-small"
                        onClick={() => deleteProduct(p.id, p.product_name)}
                      >
                        Delete
                      </button>
                    </div>
                  </td>
                </tr>
              ))}

              {rows.length === 0 && (
                <tr>
                  <td colSpan="7">
                    <EmptyState text="Hakuna bidhaa iliyopatikana." />
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function StockBadge({ stock, reorder }) {
  if (stock <= 0) return <span className="badge danger">OUT OF STOCK</span>;
  if (stock <= number(reorder))
    return <span className="badge warning">{stock}</span>;
  return <span className="badge success">{stock}</span>;
}


function InventoryHistoryPage({ businessId, refresh }) {
  const [products, setProducts] = useState([]);
  const [stockIn, setStockIn] = useState([]);
  const [sales, setSales] = useState([]);
  const [loading, setLoading] = useState(false);
  const [productFilter, setProductFilter] = useState("ALL");
  const [typeFilter, setTypeFilter] = useState("ALL");
  const [search, setSearch] = useState("");

  useEffect(() => {
    load();
  }, [businessId, refresh]);

  async function load() {
    if (!businessId) return;
    setLoading(true);
    const [{ data: p, error: pe }, { data: si, error: se }, { data: sa, error: saleError }] = await Promise.all([
      supabase.from("products").select("id, product_name, unit, opening_qty, cost_per_each, active").eq("business_id", businessId).order("product_name"),
      supabase.from("stock_in").select("id, product_id, quantity_in, cost_per_each, stock_date, staff_id").eq("business_id", businessId).order("stock_date", { ascending: true }),
      supabase.from("sales").select("id, product_id, quantity, unit_cost, selling_price, sale_date, payment_method, staff_id").eq("business_id", businessId).not("product_id", "is", null).order("sale_date", { ascending: true }),
    ]);

    if (pe || se || saleError) {
      alert((pe || se || saleError)?.message || "Imeshindikana kupakia inventory history.");
    }
    setProducts(p || []);
    setStockIn(si || []);
    setSales(sa || []);
    setLoading(false);
  }

  const productMap = useMemo(() => {
    const map = {};
    (products || []).forEach((p) => { map[p.id] = p; });
    return map;
  }, [products]);

  const movements = useMemo(() => {
    const rows = [];

    (products || []).forEach((p) => {
      rows.push({
        id: `opening-${p.id}`,
        date: null,
        product_id: p.id,
        product_name: p.product_name,
        unit: p.unit || "PCS",
        type: "OPENING",
        qty: number(p.opening_qty),
        balanceDelta: number(p.opening_qty),
        cost: number(p.cost_per_each),
        reference: "Opening Stock",
        payment: "—",
        staff_id: null,
      });
    });

    (stockIn || []).forEach((r) => {
      const p = productMap[r.product_id];
      if (!p) return;
      rows.push({
        id: `stock-${r.id}`,
        date: r.stock_date,
        product_id: r.product_id,
        product_name: p.product_name,
        unit: p.unit || "PCS",
        type: "STOCK IN",
        qty: number(r.quantity_in),
        balanceDelta: number(r.quantity_in),
        cost: number(r.cost_per_each),
        reference: "Stock In",
        payment: "—",
        staff_id: r.staff_id,
      });
    });

    (sales || []).forEach((r) => {
      const p = productMap[r.product_id];
      if (!p) return;
      rows.push({
        id: `sale-${r.id}`,
        date: r.sale_date,
        product_id: r.product_id,
        product_name: p.product_name,
        unit: p.unit || "PCS",
        type: "SALE",
        qty: number(r.quantity),
        balanceDelta: -number(r.quantity),
        cost: number(r.unit_cost),
        reference: "Sale",
        payment: r.payment_method || "—",
        staff_id: r.staff_id,
      });
    });

    const filtered = rows.filter((r) => {
      const matchesProduct = productFilter === "ALL" || r.product_id === productFilter;
      const matchesType = typeFilter === "ALL" || r.type === typeFilter;
      const q = search.trim().toLowerCase();
      const matchesSearch = !q || r.product_name.toLowerCase().includes(q) || r.type.toLowerCase().includes(q) || r.reference.toLowerCase().includes(q);
      return matchesProduct && matchesType && matchesSearch;
    });

    filtered.sort((a, b) => {
      if (a.product_id !== b.product_id) return a.product_name.localeCompare(b.product_name);
      if (!a.date && !b.date) return -1;
      if (!a.date) return -1;
      if (!b.date) return 1;
      return new Date(a.date) - new Date(b.date) || (a.type === "STOCK IN" ? -1 : 1);
    });

    const balances = {};
    return filtered.map((r) => {
      balances[r.product_id] = number(balances[r.product_id]) + r.balanceDelta;
      return { ...r, balance: balances[r.product_id] };
    });
  }, [products, stockIn, sales, productMap, productFilter, typeFilter, search]);

  const totals = useMemo(() => ({
    stockIn: movements.filter((x) => x.type === "STOCK IN").reduce((a, x) => a + x.qty, 0),
    sold: movements.filter((x) => x.type === "SALE").reduce((a, x) => a + x.qty, 0),
    opening: movements.filter((x) => x.type === "OPENING").reduce((a, x) => a + x.qty, 0),
  }), [movements]);

  function printHistory() {
    if (!movements.length) {
      alert("Hakuna inventory history ya kuchapisha.");
      return;
    }
    const w = window.open("", "_blank", "width=1100,height=800");
    if (!w) {
      alert("Browser imezuia print window. Ruhusu pop-ups kisha ujaribu tena.");
      return;
    }
    w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Inventory History</title><style>body{font-family:Arial,sans-serif;padding:24px;color:#111;font-size:12px}h1{margin:0 0 4px}.muted{color:#666;margin-bottom:18px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ccc;padding:7px}th{background:#f2f4f7;text-align:left}.num{text-align:right}.summary{margin:14px 0;font-weight:700}@media print{body{padding:0}}</style></head><body><h1>Inventory Movement History</h1><div class="muted">Generated: ${escapeHtml(formatDate(new Date()))}</div><div class="summary">Opening: ${escapeHtml(String(totals.opening))} | Stock In: ${escapeHtml(String(totals.stockIn))} | Sold: ${escapeHtml(String(totals.sold))}</div><table><thead><tr><th>Date</th><th>Product</th><th>Type</th><th>Qty</th><th>Balance</th><th>Cost</th><th>Reference</th><th>Payment</th></tr></thead><tbody>${movements.map(r => `<tr><td>${escapeHtml(r.date ? formatDate(r.date) : "Opening")}</td><td>${escapeHtml(r.product_name)}</td><td>${escapeHtml(r.type)}</td><td class="num">${escapeHtml(String(r.qty))}</td><td class="num">${escapeHtml(String(r.balance))}</td><td class="num">${escapeHtml(money(r.cost))}</td><td>${escapeHtml(r.reference)}</td><td>${escapeHtml(r.payment)}</td></tr>`).join("")}</tbody></table><script>window.onload=function(){window.print();}</script></body></html>`);
    w.document.close();
  }

  return (
    <div>
      <PageTitle title="Inventory History" subtitle="Fuatilia kila kuingia na kutoka kwa bidhaa pamoja na running stock balance." />

      <div className="stats-grid">
        <StatCard title="Opening Qty" value={number(totals.opening)} icon="◷" />
        <StatCard title="Stock In" value={number(totals.stockIn)} icon="↑" />
        <StatCard title="Sold Qty" value={number(totals.sold)} icon="↓" />
        <StatCard title="Movements" value={movements.length} icon="↕" />
      </div>

      <div className="panel">
        <div className="panel-header">
          <div>
            <h3>Stock Movement Ledger</h3>
            <span>Opening stock, Stock In na Sales kwa kila bidhaa.</span>
          </div>
          <button className="secondary-btn" onClick={printHistory}>🖨 Print</button>
        </div>

        <div className="form-grid compact">
          <Field label="Search" value={search} onChange={setSearch} placeholder="Product au movement..." />
          <SelectField label="Product" value={productFilter} onChange={setProductFilter} options={[{ value: "ALL", label: "All Products" }, ...(products || []).map(p => ({ value: p.id, label: `${p.product_name} (${p.unit || "PCS"})` }))]} />
          <SelectField label="Movement" value={typeFilter} onChange={setTypeFilter} options={[{ value: "ALL", label: "All Movements" }, { value: "OPENING", label: "Opening" }, { value: "STOCK IN", label: "Stock In" }, { value: "SALE", label: "Sale" }]} />
        </div>
      </div>

      <div className="panel">
        <div className="table-wrap">
          <table>
            <thead><tr><th>Date</th><th>Product</th><th>Movement</th><th>Qty</th><th>Balance</th><th>Cost</th><th>Reference</th><th>Payment</th></tr></thead>
            <tbody>
              {loading && <tr><td colSpan="8"><p>Inapakia inventory history...</p></td></tr>}
              {!loading && movements.map((r) => (
                <tr key={r.id}>
                  <td>{r.date ? formatDate(r.date) : <span className="muted">Opening</span>}</td>
                  <td><strong>{r.product_name}</strong><div className="muted">{r.unit}</div></td>
                  <td><span className={`badge ${r.type === "SALE" ? "danger" : r.type === "STOCK IN" ? "success" : "warning"}`}>{r.type}</span></td>
                  <td className={r.type === "SALE" ? "num" : "num"}>{r.type === "SALE" ? `-${r.qty}` : `+${r.qty}`}</td>
                  <td className="num"><strong>{r.balance}</strong></td>
                  <td className="num">{money(r.cost)}</td>
                  <td>{r.reference}</td>
                  <td>{r.payment}</td>
                </tr>
              ))}
              {!loading && !movements.length && <tr><td colSpan="8"><EmptyState text="Hakuna inventory movement inayolingana na filter." /></td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function StockInPage({ businessId, staff, onChanged }) {
  const [products, setProducts] = useState([]);
  const [records, setRecords] = useState([]);
  const [mode, setMode] = useState("existing");
  const [busy, setBusy] = useState(false);

  const [form, setForm] = useState({
    product_id: "",
    new_product_name: "",
    unit: "PCS",
    quantity_in: "",
    cost_per_each: "",
    selling_price: "",
    reorder_level: "5",
    stock_date: localDateTimeValue(),
  });

  useEffect(() => {
    load();
  }, [businessId]);

  async function load() {
    if (!businessId) return;

    const [{ data: p }, { data: r }] = await Promise.all([
      supabase.from("products").select("*").eq("business_id", businessId).eq("active", true).order("product_name"),
      supabase
        .from("stock_in")
        .select("*, products(product_name, unit)")
        .eq("business_id", businessId)
        .order("stock_date", { ascending: false })
        .limit(100),
    ]);

    setProducts(p || []);
    setRecords(r || []);

    if (!form.product_id && p?.length) {
      setForm((f) => ({ ...f, product_id: p[0].id }));
    }
  }

  function chooseStockProduct(id) {
    const p = products.find((x) => x.id === id);
    setForm((f) => ({
      ...f,
      product_id: id,
      cost_per_each: p?.cost_per_each ?? f.cost_per_each,
      selling_price: p?.sell_per_each ?? f.selling_price,
      unit: p?.unit || f.unit || "PCS",
      reorder_level: p?.reorder_level ?? f.reorder_level,
    }));
  }

  async function saveStock(e) {
    e.preventDefault();

    if (number(form.quantity_in) <= 0) {
      alert("Quantity lazima iwe zaidi ya 0.");
      return;
    }

    setBusy(true);

    let productId = form.product_id;

    if (mode === "new") {
      if (!form.new_product_name.trim()) {
        alert("Andika jina la bidhaa mpya.");
        setBusy(false);
        return;
      }

      const { data, error } = await supabase
        .from("products")
        .insert({
          business_id: businessId,
          product_name: form.new_product_name.trim(),
          unit: form.unit || "PCS",
          opening_qty: 0,
          cost_per_each: number(form.cost_per_each),
          sell_per_each: number(form.selling_price),
          reorder_level: number(form.reorder_level),
          active: true,
        })
        .select()
        .single();

      if (error) {
        alert(
          error.message.includes("duplicate")
            ? "Bidhaa hiyo tayari ipo. Chagua bidhaa iliyopo."
            : error.message
        );
        setBusy(false);
        return;
      }

      productId = data.id;
    }

    const { error } = await supabase.from("stock_in").insert({
      business_id: businessId,
      stock_date: new Date(form.stock_date).toISOString(),
      product_id: productId,
      quantity_in: number(form.quantity_in),
      cost_per_each: number(form.cost_per_each),
      staff_id: staff?.id || null,
    });

    if (error) {
      alert(error.message);
    } else {
      alert("Stock imehifadhiwa vizuri.");
      setForm({
        product_id: productId,
        new_product_name: "",
        unit: "PCS",
        quantity_in: "",
        cost_per_each: "",
        selling_price: "",
        reorder_level: "5",
        stock_date: localDateTimeValue(),
      });
      setMode("existing");
      await load();
      onChanged();
    }

    setBusy(false);
  }

  async function deleteRecord(id) {
    if (!confirm("Futa hii Stock In record?")) return;

    const { error } = await supabase
      .from("stock_in")
      .delete()
      .eq("id", id)
      .eq("business_id", businessId);

    if (error) alert(error.message);
    else {
      await load();
      onChanged();
    }
  }

  return (
    <div>
      <PageTitle
        title="Stock In"
        subtitle="Ongeza bidhaa zilizofika dukani."
      />

      <div className="mode-switch">
        <button
          className={mode === "existing" ? "selected" : ""}
          onClick={() => setMode("existing")}
        >
          📦 Existing Product
        </button>
        <button
          className={mode === "new" ? "selected" : ""}
          onClick={() => setMode("new")}
        >
          ＋ New Product
        </button>
      </div>

      <div className="panel">
        <form onSubmit={saveStock} className="form-grid">
          {mode === "existing" ? (
            <SelectField
              label="Product"
              value={form.product_id}
              onChange={(v) => {
                const p = products.find((x) => x.id === v);
                setForm({
                  ...form,
                  product_id: v,
                  cost_per_each: p?.cost_per_each || "",
                  selling_price: p?.sell_per_each || "",
                  unit: p?.unit || "PCS",
                });
              }}
              options={products.map((p) => ({
                value: p.id,
                label: `${p.product_name} (${p.unit})`,
              }))}
            />
          ) : (
            <Field
              label="New Product Name"
              value={form.new_product_name}
              onChange={(v) =>
                setForm({ ...form, new_product_name: v })
              }
              required
            />
          )}

          {mode === "new" && (
            <>
              <Field
                label="Unit"
                value={form.unit}
                onChange={(v) => setForm({ ...form, unit: v })}
              />
              <Field
                label="Selling Price"
                type="number"
                value={form.selling_price}
                onChange={(v) =>
                  setForm({ ...form, selling_price: v })
                }
              />
              <Field
                label="Reorder Level"
                type="number"
                value={form.reorder_level}
                onChange={(v) =>
                  setForm({ ...form, reorder_level: v })
                }
              />
            </>
          )}

          <Field
            label="Quantity In"
            type="number"
            value={form.quantity_in}
            onChange={(v) => setForm({ ...form, quantity_in: v })}
            required
          />

          <Field
            label="Cost Per Each"
            type="number"
            value={form.cost_per_each}
            onChange={(v) =>
              setForm({ ...form, cost_per_each: v })
            }
            required
          />

          <Field
            label="Date & Time"
            type="datetime-local"
            value={form.stock_date}
            onChange={(v) => setForm({ ...form, stock_date: v })}
          />

          <div className="form-actions full">
            <button className="primary-btn" disabled={busy}>
              {busy ? "Inahifadhi..." : "Save Stock In"}
            </button>
          </div>
        </form>
      </div>

      <div className="panel">
        <div className="panel-header">
          <div>
            <h3>Stock In History</h3>
            <span>Unaweza kufuta test/record isiyo sahihi.</span>
          </div>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Product</th>
                <th>Qty</th>
                <th>Cost</th>
                <th>Total</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {records.map((r) => (
                <tr key={r.id}>
                  <td>{formatDate(r.stock_date)}</td>
                  <td>
                    <strong>{r.products?.product_name || "Unknown"}</strong>
                  </td>
                  <td>{r.quantity_in}</td>
                  <td>{money(r.cost_per_each)}</td>
                  <td>{money(r.total_cost)}</td>
                  <td>
                    <button
                      className="danger-small"
                      onClick={() => deleteRecord(r.id)}
                    >
                      Delete
                    </button>
                  </td>
                </tr>
              ))}

              {!records.length && (
                <tr>
                  <td colSpan="6">
                    <EmptyState text="Hakuna Stock In records." />
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function SalesPage({ businessId, staff, business, onChanged }) {
  const [products, setProducts] = useState([]);
  const [services, setServices] = useState([]);
  const [sales, setSales] = useState([]);
  const [stockIn, setStockIn] = useState([]);
  const [busy, setBusy] = useState(false);
  const [receipt, setReceipt] = useState(null);
  const [availableCustomerCredit, setAvailableCustomerCredit] = useState(0);
  const [customerCreditRows, setCustomerCreditRows] = useState([]);

  const [type, setType] = useState("PRODUCT");
  const [productMode, setProductMode] = useState("existing");

  const [form, setForm] = useState({
    product_id: "",
    service_id: "",
    manual_name: "",
    quantity: "1",
    unit_cost: "",
    selling_price: "",
    payment_method: "CASH",
    customer_name: "",
    customer_phone: "",
    amount_received: "",
    credit_customer_name: "",
    credit_customer_phone: "",
    credit_paid_amount: "0",
    credit_due_date: "",
    sale_date: localDateTimeValue(),
  });

  useEffect(() => {
    load();
  }, [businessId]);

  async function load() {
    if (!businessId) return;

    const [
      { data: p },
      { data: sv },
      { data: s },
      { data: si },
    ] = await Promise.all([
      supabase.from("products").select("*").eq("business_id", businessId).eq("active", true).order("product_name"),
      supabase.from("services").select("*").eq("business_id", businessId).eq("active", true).order("service_name"),
      supabase
        .from("sales")
        .select("*, products(product_name), services(service_name)")
        .eq("business_id", businessId)
        .order("sale_date", { ascending: false })
        .limit(100),
      supabase.from("stock_in").select("*").eq("business_id", businessId),
    ]);

    setProducts(p || []);
    setServices(sv || []);
    setSales(s || []);
    setStockIn(si || []);

    if (p?.length && !form.product_id) {
      setForm((f) => ({
        ...f,
        product_id: p[0].id,
        unit_cost: p[0].cost_per_each,
        selling_price: p[0].sell_per_each,
      }));
    }

    if (sv?.length && !form.service_id) {
      setForm((f) => ({
        ...f,
        service_id: sv[0].id,
        unit_cost: sv[0].cost,
        selling_price: sv[0].selling_price,
      }));
    }
  }

  function currentStock(productId) {
    const p = products.find((x) => x.id === productId);
    if (!p) return 0;

    const incoming = stockIn
      .filter((x) => x.product_id === productId)
      .reduce((a, x) => a + number(x.quantity_in), 0);

    const sold = sales
      .filter((x) => x.product_id === productId)
      .reduce((a, x) => a + number(x.quantity), 0);

    return number(p.opening_qty) + incoming - sold;
  }

  function chooseProduct(id) {
    const p = products.find((x) => x.id === id);
    setForm({
      ...form,
      product_id: id,
      unit_cost: p?.cost_per_each || "",
      selling_price: p?.sell_per_each || "",
    });
  }

  function chooseService(id) {
    const s = services.find((x) => x.id === id);
    setForm({
      ...form,
      service_id: id,
      unit_cost: s?.cost || "",
      selling_price: s?.selling_price || "",
    });
  }

  async function makeSale(e) {
    e.preventDefault();

    const qty = number(form.quantity);
    const cost = number(form.unit_cost);
    const price = number(form.selling_price);

    if (qty <= 0) {
      alert("Quantity lazima iwe zaidi ya 0.");
      return;
    }

    if (price < 0 || cost < 0) {
      alert("Bei haiwezi kuwa chini ya 0.");
      return;
    }

    if (type === "PRODUCT" && productMode === "existing") {
      const available = currentStock(form.product_id);

      if (qty > available) {
        alert(
          `Stock haitoshi. Iliyopo ni ${available}. Umeingiza ${qty}.`
        );
        return;
      }
    }

    if (type === "PRODUCT" && productMode === "manual") {
      if (!form.manual_name.trim()) {
        alert("Andika jina la bidhaa.");
        return;
      }
    }

    const saleTotal = qty * price;
    const received = number(form.amount_received);
    const change = Math.max(0, received - saleTotal);

    if (form.payment_method === "CASH") {
      if (received <= 0) {
        alert("Weka kiasi alichopokea kutoka kwa mteja.");
        return;
      }
      if (received < saleTotal) {
        alert("Kiasi alichopokea hakiwezi kuwa chini ya jumla ya mauzo. Tumia MKOPO/CREDIT kama mteja hajalipa yote.");
        return;
      }
      if (change > 0 && !form.customer_name.trim()) {
        alert("Kuna change ya mteja. Weka jina la mteja ili change hiyo iwe Customer Credit.");
        return;
      }
    }

    if (form.payment_method === "CREDIT" && !form.credit_customer_name.trim()) {
      alert("Weka jina la mteja anayechukua kwa mkopo.");
      return;
    }

    if (form.payment_method === "CUSTOMER_CREDIT") {
      if (!form.customer_name.trim()) {
        alert("Weka jina la mteja mwenye Customer Credit.");
        return;
      }
      if (availableCustomerCredit < saleTotal) {
        alert(`Customer Credit ya ${form.customer_name} ni ${money(availableCustomerCredit)} tu. Inahitajika ${money(saleTotal)}.`);
        return;
      }
    }

    if (form.payment_method === "CREDIT") {
      const paidNow = number(form.credit_paid_amount);
      if (paidNow < 0 || paidNow > saleTotal) {
        alert("Kiasi alicholipa hakiwezi kuwa chini ya 0 au zaidi ya jumla ya sale.");
        return;
      }
      if (!form.credit_due_date) {
        alert("Chagua tarehe ya mwisho ya kulipa deni.");
        return;
      }
    }

    setBusy(true);

    const payload = {
      business_id: businessId,
      sale_date: new Date(form.sale_date).toISOString(),
      sale_type: type,
      product_id:
        type === "PRODUCT" && productMode === "existing"
          ? form.product_id
          : null,
      service_id: type === "SERVICE" ? form.service_id : null,
      quantity: qty,
      unit_cost: cost,
      selling_price: price,
      payment_method: form.payment_method,
      staff_id: staff?.id || null,
    };

    const { data, error } = await supabase
      .from("sales")
      .insert(payload)
      .select()
      .single();

    if (error) {
      alert(error.message);
      setBusy(false);
      return;
    }

    if (form.payment_method === "CREDIT") {
      const creditPayload = {
        business_id: businessId,
        sale_id: data.id,
        customer_name: form.credit_customer_name.trim(),
        customer_phone: form.credit_customer_phone.trim() || null,
        original_amount: qty * price,
        paid_amount: number(form.credit_paid_amount),
        balance: Math.max(0, qty * price - number(form.credit_paid_amount)),
        status: number(form.credit_paid_amount) >= qty * price ? "PAID" : number(form.credit_paid_amount) > 0 ? "PARTIAL" : "UNPAID",
        due_date: form.credit_due_date || null,
        product_id: type === "PRODUCT" && productMode === "existing" ? form.product_id : null,
        service_id: type === "SERVICE" ? form.service_id : null,
        created_by: staff?.id || null,
      };
      const { error: creditError } = await supabase.from("credit_transactions").insert(creditPayload);
      if (creditError) {
        alert(`Sale imehifadhiwa lakini deni halijaandikwa: ${creditError.message}`);
      }
    }

    // Cash change becomes a customer credit, so the business does not lose track of it.
    if (form.payment_method === "CASH" && change > 0) {
      const { data: creditRow, error: customerCreditError } = await supabase
        .from("customer_credits")
        .insert({
          business_id: businessId,
          customer_name: form.customer_name.trim(),
          customer_phone: form.customer_phone.trim() || null,
          original_amount: change,
          used_amount: 0,
          status: "AVAILABLE",
          source_sale_id: data.id,
          note: `Change ya sale ${data.id}`,
          created_by: staff?.id || null,
        })
        .select()
        .single();

      if (customerCreditError) {
        alert(`Sale imehifadhiwa lakini Customer Credit haijaandikwa: ${customerCreditError.message}`);
      } else {
        const { error: creditTxnError } = await supabase.from("customer_credit_transactions").insert({
          business_id: businessId,
          customer_credit_id: creditRow.id,
          transaction_type: "ADD",
          amount: change,
          sale_id: data.id,
          note: "Change kutoka sale",
          staff_id: staff?.id || null,
          created_by: staff?.id || null,
        });
        if (creditTxnError) console.error("Customer credit history error:", creditTxnError);
      }
    }

    if (form.payment_method === "CUSTOMER_CREDIT") {
      let remaining = saleTotal;
      for (const creditRow of customerCreditRows) {
        if (remaining <= 0) break;
        const available = number(creditRow.balance);
        const use = Math.min(available, remaining);
        if (use <= 0) continue;
        const newUsed = number(creditRow.used_amount) + use;
        const newBalance = Math.max(0, number(creditRow.original_amount) - newUsed);
        const newStatus = newBalance <= 0 ? "USED" : "PARTIAL";
        const { error: updateCreditError } = await supabase
          .from("customer_credits")
          .update({ used_amount: newUsed, status: newStatus, updated_at: new Date().toISOString() })
          .eq("id", creditRow.id)
          .eq("business_id", businessId);
        if (updateCreditError) {
          console.error("Customer credit update error:", updateCreditError);
          alert(`Sale imehifadhiwa lakini matumizi ya Customer Credit hayajakamilika: ${updateCreditError.message}`);
          break;
        }
        const { error: useTxnError } = await supabase.from("customer_credit_transactions").insert({
          business_id: businessId,
          customer_credit_id: creditRow.id,
          transaction_type: "USE",
          amount: use,
          sale_id: data.id,
          note: `Customer Credit imetumika kwenye sale ${data.id}`,
          staff_id: staff?.id || null,
          created_by: staff?.id || null,
        });
        if (useTxnError) console.error("Customer credit use history error:", useTxnError);
        remaining -= use;
      }
    }

    let itemName = form.manual_name;

    if (type === "PRODUCT" && productMode === "existing") {
      itemName =
        products.find((x) => x.id === form.product_id)?.product_name ||
        "Product";
    }

    if (type === "SERVICE") {
      itemName =
        services.find((x) => x.id === form.service_id)?.service_name ||
        "Service";
    }

    setReceipt({
      id: data.id,
      date: data.sale_date,
      itemName,
      type,
      quantity: qty,
      unitCost: cost,
      sellingPrice: price,
      total: qty * price,
      profit: qty * (price - cost),
      payment: form.payment_method,
      amountReceived: form.payment_method === "CASH" ? received : number(form.credit_paid_amount),
      changeAmount: form.payment_method === "CASH" ? change : 0,
      customerName:
        form.payment_method === "CREDIT"
          ? form.credit_customer_name.trim()
          : "",
      customerPhone:
        form.payment_method === "CREDIT"
          ? form.credit_customer_phone.trim()
          : "",
      creditPaid:
        form.payment_method === "CREDIT"
          ? number(form.credit_paid_amount)
          : 0,
      creditBalance:
        form.payment_method === "CREDIT"
          ? Math.max(0, qty * price - number(form.credit_paid_amount))
          : 0,
      dueDate:
        form.payment_method === "CREDIT"
          ? form.credit_due_date
          : "",
    });

    setForm({
      ...form,
      quantity: "1",
      manual_name: "",
      customer_name: "",
      customer_phone: "",
      amount_received: "",
      credit_customer_name: "",
      credit_customer_phone: "",
      credit_paid_amount: "0",
      credit_due_date: "",
      sale_date: localDateTimeValue(),
    });

    await load();
    onChanged();
    setBusy(false);
  }

  function printSalesHistory() {
    const rows = sales || [];
    if (!rows.length) {
      alert("Hakuna sales za kuchapisha.");
      return;
    }
    const totalSales = rows.reduce((a, x) => a + number(x.sales_total), 0);
    const totalProfit = rows.reduce((a, x) => a + number(x.profit), 0);
    const printWindow = window.open("", "_blank", "width=1000,height=800");
    if (!printWindow) {
      alert("Browser imezuia print window. Ruhusu pop-ups kisha ujaribu tena.");
      return;
    }
    printWindow.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Sales History</title>
      <style>
        body{font-family:Arial,sans-serif;padding:24px;color:#111;font-size:12px}
        h1{margin:0 0 4px} .muted{color:#666;margin-bottom:18px}
        table{width:100%;border-collapse:collapse}th,td{border:1px solid #ccc;padding:7px}th{background:#f2f4f7;text-align:left}
        .num{text-align:right}.summary{margin-top:16px;font-weight:700}
        @media print{body{padding:0}}
      </style></head><body>
      <h1>Sales History</h1><div class="muted">Generated: ${escapeHtml(formatDate(new Date()))}</div>
      <table><thead><tr><th>Date</th><th>Item</th><th>Type</th><th>Qty</th><th>Payment</th><th>Total</th><th>Profit</th></tr></thead>
      <tbody>${rows.map(x=>`<tr><td>${escapeHtml(formatDate(x.sale_date))}</td><td>${escapeHtml(x.products?.product_name||x.services?.service_name||"Manual Item")}</td><td>${escapeHtml(x.sale_type)}</td><td class="num">${number(x.quantity)}</td><td>${escapeHtml(x.payment_method)}</td><td class="num">${escapeHtml(money(x.sales_total))}</td><td class="num">${escapeHtml(money(x.profit))}</td></tr>`).join("")}</tbody>
      </table>
      <div class="summary">Total Sales: ${escapeHtml(money(totalSales))} &nbsp; | &nbsp; Total Profit: ${escapeHtml(money(totalProfit))}</div>
      <script>window.onload=function(){window.print();}</script></body></html>`);
    printWindow.document.close();
  }

  async function deleteSale(id) {
    if (!confirm("Futa hii sale?")) return;

    const { error } = await supabase
      .from("sales")
      .delete()
      .eq("id", id)
      .eq("business_id", businessId);

    if (error) alert(error.message);
    else {
      await load();
      onChanged();
    }
  }

  const total = number(form.quantity) * number(form.selling_price);

  return (
    <div>
      <PageTitle
        title="Sales Entry"
        subtitle="Fanya mauzo na toa receipt papo hapo."
      />

      <div className="sale-type-tabs">
        <button
          className={type === "PRODUCT" ? "active" : ""}
          onClick={() => setType("PRODUCT")}
        >
          📦 Product Sale
        </button>
        <button
          className={type === "SERVICE" ? "active" : ""}
          onClick={() => setType("SERVICE")}
        >
          ⚙ Service Sale
        </button>
      </div>

      {type === "PRODUCT" && (
        <div className="mode-switch">
          <button
            className={productMode === "existing" ? "selected" : ""}
            onClick={() => setProductMode("existing")}
          >
            Chagua Existing Product
          </button>
          <button
            className={productMode === "manual" ? "selected" : ""}
            onClick={() => setProductMode("manual")}
          >
            Andika Manual Item
          </button>
        </div>
      )}

      <div className="sale-layout">
        <div className="panel">
          <div className="panel-header">
            <div>
              <h3>New Sale</h3>
              <span>Jaza taarifa za mauzo</span>
            </div>
          </div>

          <form onSubmit={makeSale} className="form-grid">
            {type === "PRODUCT" && productMode === "existing" && (
              <SelectField
                label="Product"
                value={form.product_id}
                onChange={chooseProduct}
                options={products.map((p) => ({
                  value: p.id,
                  label: `${p.product_name} — Stock ${currentStock(
                    p.id
                  )}`,
                }))}
              />
            )}

            {type === "PRODUCT" && productMode === "manual" && (
              <>
                <Field
                  label="Item Name"
                  value={form.manual_name}
                  onChange={(v) =>
                    setForm({ ...form, manual_name: v })
                  }
                  placeholder="Mfano: Brown Envelope"
                  required
                />

                <div className="notice full">
                  ℹ️ Manual item haiathiri stock. Kwa stock tracking,
                  ongeza bidhaa kwanza kupitia Stock In.
                </div>
              </>
            )}

            {type === "SERVICE" && (
              <SelectField
                label="Service"
                value={form.service_id}
                onChange={chooseService}
                options={services.map((s) => ({
                  value: s.id,
                  label: s.service_name,
                }))}
              />
            )}

            <Field
              label="Quantity"
              type="number"
              value={form.quantity}
              onChange={(v) => setForm({ ...form, quantity: v })}
              required
            />

            <Field
              label="Unit Cost"
              type="number"
              value={form.unit_cost}
              onChange={(v) =>
                setForm({ ...form, unit_cost: v })
              }
              required
            />

            <Field
              label="Selling Price"
              type="number"
              value={form.selling_price}
              onChange={(v) =>
                setForm({ ...form, selling_price: v })
              }
              required
            />

            <Field
              label="Sale Date & Time"
              type="datetime-local"
              value={form.sale_date}
              onChange={(v) => setForm({ ...form, sale_date: v })}
              required
            />

            <SelectField
              label="Payment Method"
              value={form.payment_method}
              onChange={(v) =>
                setForm({ ...form, payment_method: v })
              }
              options={[
                { value: "CASH", label: "CASH" },
                { value: "M-PESA", label: "M-PESA" },
                { value: "AIRTEL MONEY", label: "AIRTEL MONEY" },
                { value: "TIGO PESA", label: "TIGO PESA" },
                { value: "NMB", label: "NMB" },
                { value: "BANK", label: "BANK" },
                { value: "OTHER", label: "OTHER" },
                { value: "CUSTOMER_CREDIT", label: "CUSTOMER CREDIT" },
                { value: "CREDIT", label: "MKOPO / CREDIT" },
              ]}
            />

            {form.payment_method === "CASH" && (
              <>
                <Field
                  label="Jina la Mteja (ikiwa kuna change)"
                  value={form.customer_name}
                  onChange={(v) => setForm({ ...form, customer_name: v })}
                  placeholder="Mfano: Juma"
                />
                <Field
                  label="Simu ya Mteja"
                  value={form.customer_phone}
                  onChange={(v) => setForm({ ...form, customer_phone: v })}
                  placeholder="2557XXXXXXXX"
                />
                <Field
                  label="Kiasi Alichopokea Mteja"
                  type="number"
                  value={form.amount_received}
                  onChange={(v) => setForm({ ...form, amount_received: v })}
                  min="0"
                />
                <div className="notice full">
                  💵 Jumla: <strong>{money(total)}</strong> | Amelipa: <strong>{money(form.amount_received)}</strong> | Change: <strong>{money(Math.max(0, number(form.amount_received) - total))}</strong>
                  {number(form.amount_received) > total && <span> — Change hii itawekwa moja kwa moja kwenye Customer Credit.</span>}
                </div>
              </>
            )}

            {form.payment_method === "CUSTOMER_CREDIT" && (
              <>
                <Field
                  label="Jina la Mteja Mwenye Credit"
                  value={form.customer_name}
                  onChange={(v) => setForm({ ...form, customer_name: v })}
                  placeholder="Mfano: Juma"
                  required
                />
                <Field
                  label="Simu ya Mteja"
                  value={form.customer_phone}
                  onChange={(v) => setForm({ ...form, customer_phone: v })}
                  placeholder="2557XXXXXXXX"
                />
                <div className="notice full">
                  💰 Customer Credit inayopatikana: <strong>{money(availableCustomerCredit)}</strong> | Jumla ya mauzo: <strong>{money(total)}</strong>
                  {availableCustomerCredit >= total && total > 0 ? <span> — Credit inatosha kulipia sale hii.</span> : <span> — Credit haitoshi au jina halijapatikana.</span>}
                </div>
              </>
            )}

            {form.payment_method === "CREDIT" && (
              <>
                <Field
                  label="Jina la Mteja wa Mkopo"
                  value={form.credit_customer_name}
                  onChange={(v) => setForm({ ...form, credit_customer_name: v })}
                  placeholder="Mfano: Juma"
                  required
                />
                <Field
                  label="Simu ya Mteja"
                  value={form.credit_customer_phone}
                  onChange={(v) => setForm({ ...form, credit_customer_phone: v })}
                  placeholder="2557XXXXXXXX"
                />
                <Field
                  label="Kiasi Alicholipa Sasa"
                  type="number"
                  value={form.credit_paid_amount}
                  onChange={(v) => setForm({ ...form, credit_paid_amount: v })}
                  min="0"
                />
                <Field
                  label="Deni Litalipwa Tarehe"
                  type="date"
                  value={form.credit_due_date}
                  onChange={(v) => setForm({ ...form, credit_due_date: v })}
                  required
                />
                <div className="notice full">💳 Jumla: <strong>{money(total)}</strong> | Amelipa sasa: <strong>{money(form.credit_paid_amount)}</strong> | Salio: <strong>{money(Math.max(0, total - number(form.credit_paid_amount)))}</strong></div>
              </>
            )}

            <div className="sale-total full">
              <span>Total</span>
              <strong>{money(total)}</strong>
            </div>

            <div className="form-actions full">
              <button className="primary-btn sale-save" disabled={busy}>
                {busy ? "Inahifadhi..." : "✓ COMPLETE SALE"}
              </button>
            </div>
          </form>
        </div>

        <div className="panel sale-preview">
          <div className="preview-label">SALE PREVIEW</div>
          <div className="preview-total">{money(total)}</div>
          <div className="preview-line">
            <span>Quantity</span>
            <strong>{form.quantity || 0}</strong>
          </div>
          <div className="preview-line">
            <span>Unit Price</span>
            <strong>{money(form.selling_price)}</strong>
          </div>
          <div className="preview-line">
            <span>Payment</span>
            <strong>{form.payment_method}</strong>
          </div>
          <div className="preview-hint">
            Receipt itatokea automatically baada ya sale kufanikiwa.
          </div>
        </div>
      </div>

      <div className="panel">
        <div className="panel-header">
          <div>
            <h3>Sales History</h3>
            <span>Transactions 100 za mwisho</span>
          </div>
          <div className="panel-actions">
            <button
              className="secondary-btn"
              onClick={async () => {
                const { data, error } = await supabase
                  .from("sales")
                  .select("*, products(product_name), services(service_name)")
                  .eq("business_id", businessId)
                  .order("sale_date", { ascending: false });
                if (error) {
                  alert(error.message);
                  return;
                }
                const rows = (data || []).map((x) => ({
                  Date: formatDate(x.sale_date),
                  Item: x.products?.product_name || x.services?.service_name || "Manual Item",
                  Type: x.sale_type,
                  Quantity: x.quantity,
                  "Unit Cost": x.unit_cost,
                  "Selling Price": x.selling_price,
                  Total: x.sales_total,
                  Profit: x.profit,
                  "Payment Method": x.payment_method,
                }));
                exportToExcel(rows, `Bless-Stationery-Sales-${new Date().toISOString().slice(0, 10)}.xlsx`, "Sales");
              }}
            >
              ⬇ Export Excel
            </button>
            <button
              className="secondary-btn"
              onClick={printSalesHistory}
            >
              🖨 Print History
            </button>
          </div>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Item</th>
                <th>Type</th>
                <th>Qty</th>
                <th>Payment</th>
                <th>Total</th>
                <th>Profit</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {sales.map((s) => (
                <tr key={s.id}>
                  <td>{formatDate(s.sale_date)}</td>
                  <td>
                    <strong>
                      {s.products?.product_name ||
                        s.services?.service_name ||
                        "Manual Item"}
                    </strong>
                  </td>
                  <td>{s.sale_type}</td>
                  <td>{s.quantity}</td>
                  <td>
                    <span className="payment-badge">
                      {s.payment_method}
                    </span>
                  </td>
                  <td>{money(s.sales_total)}</td>
                  <td className="profit-text">{money(s.profit)}</td>
                  <td>
                    <button
                      className="danger-small"
                      onClick={() => deleteSale(s.id)}
                    >
                      Delete
                    </button>
                  </td>
                </tr>
              ))}

              {!sales.length && (
                <tr>
                  <td colSpan="8">
                    <EmptyState text="Hakuna sales bado." />
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {receipt && (
        <ReceiptModal
          receipt={receipt}
          business={business}
          onClose={() => setReceipt(null)}
        />
      )}
    </div>
  );
}

function ReceiptModal({ receipt, business, onClose }) {
  function printReceipt() {
    const printWindow = window.open("", "_blank", "width=420,height=760");

    if (!printWindow) {
      alert("Browser imezuia dirisha la print. Ruhusu pop-ups kisha ujaribu tena.");
      return;
    }

    const escapeHtml = (value) =>
      String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");

    const businessName = business?.business_name || "Bless Stationery";
    const businessPhone = business?.phone || "";
    const businessAddress = business?.address || "";
    const receiptNo = String(receipt.id || "").slice(-8).toUpperCase();
    const creditSection =
      receipt.payment === "CREDIT"
        ? `
          <div class="divider"></div>
          <div class="row"><span>Customer</span><strong>${escapeHtml(receipt.customerName || "-")}</strong></div>
          ${
            receipt.customerPhone
              ? `<div class="row"><span>Phone</span><strong>${escapeHtml(receipt.customerPhone)}</strong></div>`
              : ""
          }
          <div class="row"><span>Paid Now</span><strong>${escapeHtml(money(receipt.creditPaid))}</strong></div>
          <div class="row balance"><span>Balance</span><strong>${escapeHtml(money(receipt.creditBalance))}</strong></div>
          ${
            receipt.dueDate
              ? `<div class="row"><span>Due Date</span><strong>${escapeHtml(receipt.dueDate)}</strong></div>`
              : ""
          }
        `
        : "";

    printWindow.document.write(`<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>Receipt ${escapeHtml(receiptNo)}</title>
<style>
  @page { size: 80mm auto; margin: 4mm; }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: Arial, sans-serif;
    color: #111;
    font-size: 12px;
    background: #fff;
  }
  .receipt {
    width: 72mm;
    margin: 0 auto;
  }
  .center { text-align: center; }
  h1 { margin: 0 0 3px; font-size: 19px; }
  .muted { color: #555; }
  .small { font-size: 10px; }
  .title { font-weight: 700; letter-spacing: .5px; margin-top: 4px; }
  .divider {
    border-top: 1px dashed #555;
    margin: 9px 0;
  }
  .row {
    display: flex;
    justify-content: space-between;
    gap: 10px;
    margin: 5px 0;
  }
  .row span { color: #555; }
  .row strong { text-align: right; }
  .item {
    display: flex;
    justify-content: space-between;
    gap: 10px;
    margin: 7px 0;
  }
  .item-name { max-width: 45mm; font-weight: 700; }
  .total {
    display: flex;
    justify-content: space-between;
    font-size: 16px;
    font-weight: 700;
    margin: 9px 0;
  }
  .balance strong { font-size: 14px; }
  .footer { text-align: center; margin-top: 14px; }
  @media print {
    body { width: 72mm; }
  }
</style>
</head>
<body>
  <div class="receipt">
    <div class="center">
      <h1>${escapeHtml(businessName)}</h1>
      ${businessAddress ? `<div class="muted small">${escapeHtml(businessAddress)}</div>` : ""}
      ${businessPhone ? `<div class="muted small">${escapeHtml(businessPhone)}</div>` : ""}
      <div class="title">SALES RECEIPT</div>
    </div>

    <div class="divider"></div>

    <div class="row">
      <span>Receipt No.</span>
      <strong>${escapeHtml(receiptNo)}</strong>
    </div>
    <div class="row">
      <span>Date</span>
      <strong>${escapeHtml(formatDate(receipt.date))}</strong>
    </div>

    <div class="divider"></div>

    <div class="item">
      <span class="item-name">${escapeHtml(receipt.itemName)}</span>
      <strong>${escapeHtml(money(receipt.total))}</strong>
    </div>
    <div class="row">
      <span>${escapeHtml(receipt.quantity)} × ${escapeHtml(money(receipt.sellingPrice))}</span>
      <strong>${escapeHtml(money(receipt.total))}</strong>
    </div>

    <div class="divider"></div>

    <div class="total">
      <span>TOTAL</span>
      <span>${escapeHtml(money(receipt.total))}</span>
    </div>

    <div class="row">
      <span>Payment</span>
      <strong>${escapeHtml(receipt.payment)}</strong>
    </div>

    ${creditSection}

    <div class="footer">
      <div>Asante kwa kufanya biashara nasi.</div>
      <div class="muted small">${escapeHtml(businessName)}</div>
    </div>
  </div>
<script>
  window.onload = function () {
    window.focus();
    window.print();
  };
  window.onafterprint = function () {
    window.close();
  };
</script>
</body>
</html>`);

    printWindow.document.close();
  }

  return (
    <div className="modal-overlay">
      <div className="receipt-modal">
        <div className="receipt-actions no-print">
          <button className="secondary-btn" onClick={onClose}>
            Close
          </button>
          <button className="primary-btn" onClick={printReceipt}>
            🖨 Print Receipt
          </button>
        </div>

        <div className="receipt-print">
          <div className="receipt-header">
            <div className="receipt-logo">B</div>
            <h2>{business?.business_name || "Bless Stationery"}</h2>
            <p>SALES RECEIPT</p>
          </div>

          <div className="receipt-info">
            <div>
              <span>Receipt No.</span>
              <strong>{receipt.id.slice(-8).toUpperCase()}</strong>
            </div>
            <div>
              <span>Date</span>
              <strong>{formatDate(receipt.date)}</strong>
            </div>
          </div>

          <div className="receipt-divider" />

          <div className="receipt-item">
            <strong>{receipt.itemName}</strong>
            <span>
              {receipt.quantity} × {money(receipt.sellingPrice)}
            </span>
          </div>

          <div className="receipt-divider" />

          <div className="receipt-total">
            <span>TOTAL</span>
            <strong>{money(receipt.total)}</strong>
          </div>

          <div className="receipt-info">
            <div>
              <span>Payment</span>
              <strong>{receipt.payment}</strong>
            </div>
            <div>
              <span>Profit</span>
              <strong>{money(receipt.profit)}</strong>
            </div>
          </div>

          <div className="receipt-footer">
            <p>Asante kwa kufanya biashara nasi.</p>
            <small>{business?.business_name || "Bless Stationery"}</small>
          </div>
        </div>

        <div className="print-note no-print">
          Ukibonyeza Print, Chrome/Windows itakupa sehemu ya kuchagua
          printer.
        </div>
      </div>
    </div>
  );
}

function ServicesPage({ businessId, onChanged }) {
  const [services, setServices] = useState([]);
  const [show, setShow] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [search, setSearch] = useState("");
  const [form, setForm] = useState({
    service_name: "",
    selling_price: "",
    cost: "0",
  });

  useEffect(() => {
    load();
  }, [businessId]);

  async function load() {
    if (!businessId) return;
    const { data } = await supabase
      .from("services")
      .select("*")
      .eq("business_id", businessId)
      .order("service_name");

    setServices(data || []);
  }

  function resetServiceForm() {
    setForm({ service_name: "", selling_price: "", cost: "0" });
    setEditingId(null);
  }

  function startEditService(row) {
    setEditingId(row.id);
    setForm({
      service_name: row.service_name || "",
      selling_price: String(row.selling_price ?? ""),
      cost: String(row.cost ?? "0"),
    });
    setShow(true);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function addService(e) {
    e.preventDefault();
    if (!form.service_name.trim()) { alert("Weka jina la service."); return; }
    if (positiveNumber(form.selling_price) < 0 || positiveNumber(form.cost) < 0) { alert("Bei haiwezi kuwa chini ya 0."); return; }

    const payload = {
      service_name: form.service_name.trim(),
      selling_price: positiveNumber(form.selling_price),
      cost: positiveNumber(form.cost),
      active: true,
    };

    const result = editingId
      ? await supabase.from("services").update(payload).eq("id", editingId).eq("business_id", businessId)
      : await supabase.from("services").insert({ business_id: businessId, ...payload });

    if (result.error) {
      alert(result.error.message);
    } else {
      resetServiceForm();
      setShow(false);
      await load();
      onChanged();
    }
  }

  async function deleteService(id) {
    if (!confirm("Futa service hii?")) return;

    const { error } = await supabase
      .from("services")
      .delete()
      .eq("id", id)
      .eq("business_id", businessId);

    if (error) {
      alert(
        "Service hii haiwezi kufutwa kwa sababu tayari imetumika kwenye sales."
      );
    } else {
      await load();
      onChanged();
    }
  }

  return (
    <div>
      <PageTitle
        title="Services"
        subtitle="Simamia huduma na bei zake."
        button={show ? "Close Form" : "+ Add Service"}
        onClick={() => { if (show) { setShow(false); resetServiceForm(); } else setShow(true); }}
      />

      {show && (
        <div className="panel">
          <div className="panel-header"><div><h3>{editingId ? "Hariri Service" : "Ongeza Service"}</h3><span>Bei ya kuuza na gharama ya huduma.</span></div></div>
          <form onSubmit={addService} className="form-grid">
            <Field
              label="Service Name"
              value={form.service_name}
              onChange={(v) =>
                setForm({ ...form, service_name: v })
              }
              required
            />
            <Field
              label="Selling Price"
              type="number"
              value={form.selling_price}
              onChange={(v) =>
                setForm({ ...form, selling_price: v })
              }
              required
            />
            <Field
              label="Cost"
              type="number"
              value={form.cost}
              onChange={(v) => setForm({ ...form, cost: v })}
            />

            <div className="form-actions full">
              <button className="primary-btn">{editingId ? "Update Service" : "Save Service"}</button>
            </div>
          </form>
        </div>
      )}

      <div className="panel">
        <div className="toolbar">
          <div className="search-box"><span>🔎</span><input placeholder="Search service..." value={search} onChange={e=>setSearch(e.target.value)} /></div>
          <button className="secondary-btn small-btn" onClick={() => exportToExcel(
            services.map(s => ({ Service: s.service_name, "Selling Price": s.selling_price, Cost: s.cost, Profit: number(s.selling_price) - number(s.cost) })),
            `Bless-Stationery-Services-${todayDateInput()}.xlsx`, "Services"
          )}>⬇ Export</button>
          <span className="record-count">{services.filter(s => s.service_name.toLowerCase().includes(search.toLowerCase())).length} services</span>
        </div>
      </div>

      <div className="service-grid">
        {services.filter((s) => s.service_name.toLowerCase().includes(search.toLowerCase())).map((s) => (
          <div className="service-card" key={s.id}>
            <div className="service-icon">⚙</div>
            <h3>{s.service_name}</h3>
            <span>Price</span>
            <strong>{money(s.selling_price)}</strong>
            <div className="row-actions">
              <button className="secondary-btn small-btn" onClick={() => startEditService(s)}>
                Edit
              </button>
              <button className="danger-small" onClick={() => deleteService(s.id)}>
                Delete
              </button>
            </div>
          </div>
        ))}

        {!services.filter((s) => s.service_name.toLowerCase().includes(search.toLowerCase())).length && (
          <EmptyState text={services.length ? "Hakuna service inayolingana na search." : "Hakuna services."} />
        )}
      </div>
    </div>
  );
}

function ExpensesPage({ businessId, staff, onChanged }) {
  const [records, setRecords] = useState([]);
  const [busy, setBusy] = useState(false);
  const [form, setForm] = useState({
    expense_item: "",
    category: "Other",
    amount: "",
    description: "",
    expense_date: localDateTimeValue(),
  });

  useEffect(() => {
    load();
  }, [businessId]);

  async function load() {
    if (!businessId) return;
    const { data } = await supabase
      .from("expenses")
      .select("*")
      .eq("business_id", businessId)
      .order("expense_date", { ascending: false })
      .limit(200);

    setRecords(data || []);
  }

  async function save(e) {
    e.preventDefault();

    if (!form.expense_item.trim()) { alert("Weka jina la expense."); return; }
    if (number(form.amount) <= 0) { alert("Amount lazima iwe zaidi ya 0."); return; }

    setBusy(true);
    const { error } = await supabase.from("expenses").insert({
      business_id: businessId,
      expense_date: new Date(form.expense_date).toISOString(),
      expense_item: form.expense_item.trim(),
      category: form.category,
      amount: number(form.amount),
      description: form.description,
      staff_id: staff?.id || null,
    });

    if (error) alert(error.message);
    else {
      setForm({
        expense_item: "",
        category: "Other",
        amount: "",
        description: "",
        expense_date: localDateTimeValue(),
      });
      await load();
      onChanged();
    }
    setBusy(false);
  }

  async function remove(id) {
    if (!confirm("Futa expense hii?")) return;

    const { error } = await supabase
      .from("expenses")
      .delete()
      .eq("id", id)
      .eq("business_id", businessId);

    if (error) alert(error.message);
    else {
      await load();
      onChanged();
    }
  }

  return (
    <div>
      <PageTitle
        title="Expenses"
        subtitle="Rekodi matumizi ya biashara."
      />

      <div className="panel">
        <form onSubmit={save} className="form-grid">
          <Field
            label="Expense Item"
            value={form.expense_item}
            onChange={(v) => setForm({ ...form, expense_item: v })}
            required
          />

          <SelectField
            label="Category"
            value={form.category}
            onChange={(v) => setForm({ ...form, category: v })}
            options={[
              { value: "Rent", label: "Rent" },
              { value: "Transport", label: "Transport" },
              { value: "Electricity", label: "Electricity" },
              { value: "Internet", label: "Internet" },
              { value: "Salary", label: "Salary" },
              { value: "Supplies", label: "Supplies" },
              { value: "Other", label: "Other" },
            ]}
          />

          <Field
            label="Expense Date & Time"
            type="datetime-local"
            value={form.expense_date}
            onChange={(v) => setForm({ ...form, expense_date: v })}
            required
          />

          <Field
            label="Amount"
            type="number"
            value={form.amount}
            onChange={(v) => setForm({ ...form, amount: v })}
            required
          />

          <Field
            label="Description"
            value={form.description}
            onChange={(v) =>
              setForm({ ...form, description: v })
            }
          />

          <div className="form-actions full">
            <button className="primary-btn" disabled={busy}>{busy ? "Inahifadhi..." : "Save Expense"}</button>
          </div>
        </form>
      </div>

      <div className="stats-grid small-stats"><StatCard title="Total Expenses" value={money(records.reduce((a,r)=>a+number(r.amount),0))} icon="💸" tone="red" /><StatCard title="Records" value={records.length} icon="🧾" tone="blue" /></div>
      <div className="panel">
        <div className="panel-header">
          <div>
            <h3>Expense History</h3>
            <span>{records.length} records</span>
          </div>
          <button
            className="secondary-btn"
            onClick={() => {
              const rows = records.map((r) => ({
                Date: formatDate(r.expense_date),
                Item: r.expense_item,
                Category: r.category,
                Description: r.description || "",
                Amount: r.amount,
              }));
              exportToExcel(rows, `Bless-Stationery-Expenses-${new Date().toISOString().slice(0, 10)}.xlsx`, "Expenses");
            }}
          >
            ⬇ Export Excel
          </button>
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Item</th>
                <th>Category</th>
                <th>Description</th>
                <th>Amount</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {records.map((r) => (
                <tr key={r.id}>
                  <td>{formatDate(r.expense_date)}</td>
                  <td>
                    <strong>{r.expense_item}</strong>
                  </td>
                  <td>{r.category}</td>
                  <td>{r.description || "-"}</td>
                  <td>{money(r.amount)}</td>
                  <td>
                    <button
                      className="danger-small"
                      onClick={() => remove(r.id)}
                    >
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function DepositsPage({ businessId, staff, onChanged }) {
  const [records, setRecords] = useState([]);
  const [busy, setBusy] = useState(false);
  const [form, setForm] = useState({
    amount: "",
    method: "CASH",
    depositor: "",
    reference_no: "",
    note: "",
    deposit_date: localDateTimeValue(),
  });

  useEffect(() => {
    load();
  }, [businessId]);

  async function load() {
    if (!businessId) return;
    const { data } = await supabase
      .from("deposits")
      .select("*")
      .eq("business_id", businessId)
      .order("deposit_date", { ascending: false })
      .limit(200);

    setRecords(data || []);
  }

  async function save(e) {
    e.preventDefault();
    if (number(form.amount) <= 0) { alert("Amount lazima iwe zaidi ya 0."); return; }

    setBusy(true);
    const { error } = await supabase.from("deposits").insert({
      business_id: businessId,
      deposit_date: new Date(form.deposit_date).toISOString(),
      amount: number(form.amount),
      method: form.method,
      depositor: form.depositor,
      reference_no: form.reference_no,
      note: form.note,
      staff_id: staff?.id || null,
    });

    if (error) alert(error.message);
    else {
      setForm({
        amount: "",
        method: "CASH",
        depositor: "",
        reference_no: "",
        note: "",
        deposit_date: localDateTimeValue(),
      });
      await load();
      onChanged();
    }
    setBusy(false);
  }

  async function remove(id) {
    if (!confirm("Futa deposit hii?")) return;

    const { error } = await supabase
      .from("deposits")
      .delete()
      .eq("id", id)
      .eq("business_id", businessId);

    if (error) alert(error.message);
    else {
      await load();
      onChanged();
    }
  }

  return (
    <div>
      <PageTitle
        title="Deposits"
        subtitle="Rekodi fedha zilizowekwa/deposited."
      />

      <div className="panel">
        <form onSubmit={save} className="form-grid">
          <Field
            label="Amount"
            type="number"
            value={form.amount}
            onChange={(v) => setForm({ ...form, amount: v })}
            required
          />

          <SelectField
            label="Method"
            value={form.method}
            onChange={(v) => setForm({ ...form, method: v })}
            options={[
              { value: "CASH", label: "CASH" },
              { value: "M-PESA", label: "M-PESA" },
              { value: "AIRTEL MONEY", label: "AIRTEL MONEY" },
              { value: "TIGO PESA", label: "TIGO PESA" },
              { value: "NMB", label: "NMB" },
              { value: "BANK", label: "BANK" },
              { value: "OTHER", label: "OTHER" },
            ]}
          />

          <Field
            label="Deposit Date & Time"
            type="datetime-local"
            value={form.deposit_date}
            onChange={(v) => setForm({ ...form, deposit_date: v })}
            required
          />

          <Field
            label="Depositor"
            value={form.depositor}
            onChange={(v) => setForm({ ...form, depositor: v })}
          />

          <Field
            label="Reference No."
            value={form.reference_no}
            onChange={(v) =>
              setForm({ ...form, reference_no: v })
            }
          />

          <Field
            label="Note"
            value={form.note}
            onChange={(v) => setForm({ ...form, note: v })}
          />

          <div className="form-actions full">
            <button className="primary-btn" disabled={busy}>{busy ? "Inahifadhi..." : "Save Deposit"}</button>
          </div>
        </form>
      </div>

      <div className="stats-grid small-stats"><StatCard title="Total Deposits" value={money(records.reduce((a,r)=>a+number(r.amount),0))} icon="🏦" tone="green" /><StatCard title="Records" value={records.length} icon="🧾" tone="blue" /></div>
      <div className="panel">
        <div className="panel-header">
          <div>
            <h3>Deposit History</h3>
            <span>{records.length} records</span>
          </div>
          <button
            className="secondary-btn"
            onClick={() => {
              const rows = records.map((r) => ({
                Date: formatDate(r.deposit_date),
                Amount: r.amount,
                Method: r.method,
                Depositor: r.depositor || "",
                Reference: r.reference_no || "",
                Note: r.note || "",
              }));
              exportToExcel(rows, `Bless-Stationery-Deposits-${new Date().toISOString().slice(0, 10)}.xlsx`, "Deposits");
            }}
          >
            ⬇ Export Excel
          </button>
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Amount</th>
                <th>Method</th>
                <th>Depositor</th>
                <th>Reference</th>
                <th>Note</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {records.map((r) => (
                <tr key={r.id}>
                  <td>{formatDate(r.deposit_date)}</td>
                  <td>{money(r.amount)}</td>
                  <td>{r.method}</td>
                  <td>{r.depositor || "-"}</td>
                  <td>{r.reference_no || "-"}</td>
                  <td>{r.note || "-"}</td>
                  <td>
                    <button
                      className="danger-small"
                      onClick={() => remove(r.id)}
                    >
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function ReportsPage({ businessId, refresh }) {
  const [period, setPeriod] = useState("today");
  const [data, setData] = useState({ sales: [], expenses: [], deposits: [], credits: [], attendance: [] });
  const [showManual, setShowManual] = useState(false);
  const [manualFrom, setManualFrom] = useState(() => new Date().toISOString().slice(0, 10));
  const [manualTo, setManualTo] = useState(() => new Date().toISOString().slice(0, 10));
  const [manualType, setManualType] = useState("all");
  const [manualTitle, setManualTitle] = useState("Manual Business Report");
  const [manualNote, setManualNote] = useState("");
  const [manualData, setManualData] = useState(null);
  const [manualBusy, setManualBusy] = useState(false);

  useEffect(() => {
    load();
  }, [period, refresh, businessId]);

  async function load() {
    if (!businessId) return;
    const start = period === "today" ? todayStart() : monthStart();

    const [
      { data: sales },
      { data: expenses },
      { data: deposits },
      { data: credits },
      { data: attendance },
    ] = await Promise.all([
      supabase
        .from("sales")
        .select("*")
        .eq("business_id", businessId)
        .gte("sale_date", start)
        .order("sale_date", { ascending: false }),
      supabase
        .from("expenses")
        .select("*")
        .eq("business_id", businessId)
        .gte("expense_date", start)
        .order("expense_date", { ascending: false }),
      supabase
        .from("deposits")
        .select("*")
        .eq("business_id", businessId)
        .gte("deposit_date", start)
        .order("deposit_date", { ascending: false }),
      supabase
        .from("credit_transactions")
        .select("*")
        .eq("business_id", businessId)
        .gte("created_at", start)
        .order("created_at", { ascending: false }),
      supabase
        .from("staff_attendance")
        .select("*, staff(staff_name)")
        .eq("business_id", businessId)
        .gte("attendance_date", start.slice(0,10))
        .order("check_in_at", { ascending: false }),
    ]);

    setData({
      sales: sales || [],
      expenses: expenses || [],
      deposits: deposits || [],
      credits: credits || [],
      attendance: attendance || [],
    });
  }

  async function generateManualReport() {
    if (!businessId) return;
    if (!manualFrom || !manualTo) {
      alert("Chagua tarehe ya kuanzia na tarehe ya mwisho.");
      return;
    }
    if (manualFrom > manualTo) {
      alert("Tarehe ya kuanzia haiwezi kuwa baada ya tarehe ya mwisho.");
      return;
    }

    setManualBusy(true);
    const from = new Date(`${manualFrom}T00:00:00`).toISOString();
    const to = new Date(`${manualTo}T23:59:59.999`).toISOString();

    const requests = [];
    if (manualType === "all" || manualType === "sales") {
      requests.push(
        supabase
          .from("sales")
          .select("*")
          .eq("business_id", businessId)
          .gte("sale_date", from)
          .lte("sale_date", to)
          .order("sale_date", { ascending: false })
      );
    } else {
      requests.push(Promise.resolve({ data: [], error: null }));
    }

    if (manualType === "all" || manualType === "expenses") {
      requests.push(
        supabase
          .from("expenses")
          .select("*")
          .eq("business_id", businessId)
          .gte("expense_date", from)
          .lte("expense_date", to)
          .order("expense_date", { ascending: false })
      );
    } else {
      requests.push(Promise.resolve({ data: [], error: null }));
    }

    if (manualType === "all" || manualType === "deposits") {
      requests.push(
        supabase
          .from("deposits")
          .select("*")
          .eq("business_id", businessId)
          .gte("deposit_date", from)
          .lte("deposit_date", to)
          .order("deposit_date", { ascending: false })
      );
    } else {
      requests.push(Promise.resolve({ data: [], error: null }));
    }

    const [salesResult, expensesResult, depositsResult] = await Promise.all(requests);
    const firstError = salesResult.error || expensesResult.error || depositsResult.error;

    if (firstError) {
      alert(firstError.message);
      setManualBusy(false);
      return;
    }

    const result = {
      sales: salesResult.data || [],
      expenses: expensesResult.data || [],
      deposits: depositsResult.data || [],
    };

    setManualData(result);
    setManualBusy(false);
  }

  function printReport(report = manualData, title = manualTitle) {
    if (!report) {
      alert("Kwanza tengeneza Manual Report.");
      return;
    }

    const printWindow = window.open("", "_blank", "width=1000,height=800");
    if (!printWindow) {
      alert("Browser imezuia print window. Ruhusu pop-ups kisha ujaribu tena.");
      return;
    }

    const salesTotal = report.sales.reduce((a, x) => a + number(x.sales_total), 0);
    const profit = report.sales.reduce((a, x) => a + number(x.profit), 0);
    const expensesTotal = report.expenses.reduce((a, x) => a + number(x.amount), 0);
    const depositsTotal = report.deposits.reduce((a, x) => a + number(x.amount), 0);

    const salesRows = report.sales.map((x) => `
      <tr><td>${escapeHtml(formatDate(x.sale_date))}</td><td>${escapeHtml(x.sale_type || "-")}</td><td>${number(x.quantity)}</td><td>${escapeHtml(money(x.sales_total))}</td><td>${escapeHtml(money(x.profit))}</td><td>${escapeHtml(x.payment_method || "-")}</td></tr>
    `).join("");
    const expenseRows = report.expenses.map((x) => `
      <tr><td>${escapeHtml(formatDate(x.expense_date))}</td><td>${escapeHtml(x.expense_item || "-")}</td><td>${escapeHtml(x.category || "-")}</td><td>${escapeHtml(money(x.amount))}</td><td>${escapeHtml(x.description || "-")}</td></tr>
    `).join("");
    const depositRows = report.deposits.map((x) => `
      <tr><td>${escapeHtml(formatDate(x.deposit_date))}</td><td>${escapeHtml(money(x.amount))}</td><td>${escapeHtml(x.method || "-")}</td><td>${escapeHtml(x.depositor || "-")}</td><td>${escapeHtml(x.reference_no || "-")}</td></tr>
    `).join("");

    printWindow.document.write(`<!doctype html><html><head><title>${title}</title><style>
      body{font-family:Arial,sans-serif;color:#111;padding:30px;font-size:12px}h1{margin:0 0 6px}h2{margin-top:28px}.muted{color:#666}.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:20px 0}.box{border:1px solid #ddd;padding:12px;border-radius:6px}.box b{display:block;font-size:16px;margin-top:5px}table{width:100%;border-collapse:collapse;margin-top:10px}th,td{border:1px solid #ddd;padding:7px;text-align:left}th{background:#f3f3f3}.note{margin:15px 0;padding:10px;border-left:4px solid #555;background:#f7f7f7}@media print{body{padding:0}.no-print{display:none}}
    </style></head><body>
      <h1>${title}</h1>
      <div class="muted">Bless Stationery Management System</div>
      <div class="muted">Kipindi: ${manualFrom} hadi ${manualTo} | Imetengenezwa: ${formatDate(new Date().toISOString())}</div>
      ${manualNote ? `<div class="note">${manualNote}</div>` : ""}
      <div class="summary"><div class="box">Sales<b>${money(salesTotal)}</b></div><div class="box">Profit<b>${money(profit)}</b></div><div class="box">Expenses<b>${money(expensesTotal)}</b></div><div class="box">Deposits<b>${money(depositsTotal)}</b></div></div>
      ${report.sales.length ? `<h2>Sales</h2><table><thead><tr><th>Date</th><th>Type</th><th>Qty</th><th>Total</th><th>Profit</th><th>Payment</th></tr></thead><tbody>${salesRows}</tbody></table>` : ""}
      ${report.expenses.length ? `<h2>Expenses</h2><table><thead><tr><th>Date</th><th>Item</th><th>Category</th><th>Amount</th><th>Description</th></tr></thead><tbody>${expenseRows}</tbody></table>` : ""}
      ${report.deposits.length ? `<h2>Deposits</h2><table><thead><tr><th>Date</th><th>Amount</th><th>Method</th><th>Depositor</th><th>Reference</th></tr></thead><tbody>${depositRows}</tbody></table>` : ""}
      <script>window.onload=function(){window.print();}</script>
    </body></html>`);
    printWindow.document.close();
    setTimeout(() => {
      try {
        printWindow.focus();
        printWindow.print();
      } catch (e) {}
    }, 300);
  }

  const salesTotal = data.sales.reduce((a, x) => a + number(x.sales_total), 0);
  const profit = data.sales.reduce((a, x) => a + number(x.profit), 0);
  const expensesTotal = data.expenses.reduce((a, x) => a + number(x.amount), 0);
  const depositsTotal = data.deposits.reduce((a, x) => a + number(x.amount), 0);
  const creditBalance = data.credits.reduce((a, x) => a + number(x.balance), 0);

  return (
    <div>
      <PageTitle
        title="Reports"
        subtitle="Taarifa za biashara kwa siku au mwezi."
      />

      <div className="report-toolbar">
        <div className="report-tabs">
          <button className={period === "today" ? "active" : ""} onClick={() => setPeriod("today")}>Leo</button>
          <button className={period === "month" ? "active" : ""} onClick={() => setPeriod("month")}>Mwezi Huu</button>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button className="secondary-btn" onClick={() => setShowManual(!showManual)}>
            📝 Manual Report
          </button>
          <button className="secondary-btn" onClick={() => printReport(data, `Bless Stationery - ${period === "today" ? "Daily" : "Monthly"} Report`)}>
            🖨 Print Report
          </button>
          <button
            className="secondary-btn"
            onClick={() => {
              if (!data.sales.length && !data.expenses.length && !data.deposits.length) {
                alert("Hakuna report data ya ku-export kwa kipindi hiki.");
                return;
              }
              const workbook = XLSX.utils.book_new();
              const salesRows = data.sales.map((x) => ({ Date: formatDate(x.sale_date), Type: x.sale_type, Quantity: x.quantity, "Unit Cost": x.unit_cost, "Selling Price": x.selling_price, Total: x.sales_total, Profit: x.profit, "Payment Method": x.payment_method }));
              const expenseRows = data.expenses.map((x) => ({ Date: formatDate(x.expense_date), Item: x.expense_item, Category: x.category, Amount: x.amount, Description: x.description || "" }));
              const depositRows = data.deposits.map((x) => ({ Date: formatDate(x.deposit_date), Amount: x.amount, Method: x.method, Depositor: x.depositor || "", Reference: x.reference_no || "", Note: x.note || "" }));
              XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(salesRows), "Sales");
              XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(expenseRows), "Expenses");
              XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(depositRows), "Deposits");
              XLSX.writeFile(workbook, `Bless-Stationery-${period === "today" ? "Daily" : "Monthly"}-Report-${new Date().toISOString().slice(0, 10)}.xlsx`);
            }}
          >
            ⬇ Export Excel
          </button>
        </div>
      </div>

      {showManual && (
        <div className="panel" style={{ marginBottom: 18 }}>
          <div className="panel-header">
            <div>
              <h3>Manual Report</h3>
              <span>Chagua kipindi na aina ya taarifa unayotaka kutoa.</span>
            </div>
          </div>
          <div className="form-grid">
            <div><label>Tarehe Kuanzia</label><input type="date" value={manualFrom} onChange={(e) => setManualFrom(e.target.value)} /></div>
            <div><label>Tarehe Mpaka</label><input type="date" value={manualTo} onChange={(e) => setManualTo(e.target.value)} /></div>
            <div><label>Aina ya Report</label><select value={manualType} onChange={(e) => setManualType(e.target.value)}><option value="all">Zote</option><option value="sales">Sales</option><option value="expenses">Expenses</option><option value="deposits">Deposits</option></select></div>
            <div><label>Kichwa cha Report</label><input type="text" value={manualTitle} onChange={(e) => setManualTitle(e.target.value)} placeholder="Mfano: Report ya Wiki" /></div>
          </div>
          <label>Maelezo / Note (optional)</label>
          <textarea value={manualNote} onChange={(e) => setManualNote(e.target.value)} placeholder="Andika maelezo ya report..." rows="3" />
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 12 }}>
            <button className="primary-btn" onClick={generateManualReport} disabled={manualBusy}>{manualBusy ? "Inatengeneza..." : "Tengeneza Manual Report"}</button>
            {manualData && <button className="secondary-btn" onClick={() => printReport(manualData, manualTitle || "Manual Business Report")}>🖨 Print Manual Report</button>}
          </div>

          {manualData && (
            <div style={{ marginTop: 18 }}>
              <div className="stats-grid">
                <StatCard title="Sales" value={money(manualData.sales.reduce((a, x) => a + number(x.sales_total), 0))} icon="💰" tone="blue" />
                <StatCard title="Profit" value={money(manualData.sales.reduce((a, x) => a + number(x.profit), 0))} icon="📈" tone="green" />
                <StatCard title="Expenses" value={money(manualData.expenses.reduce((a, x) => a + number(x.amount), 0))} icon="💸" tone="red" />
                <StatCard title="Deposits" value={money(manualData.deposits.reduce((a, x) => a + number(x.amount), 0))} icon="🏦" tone="purple" />
              </div>
              <p className="muted">Report imetengenezwa kwa tarehe {manualFrom} hadi {manualTo}.</p>
            </div>
          )}
        </div>
      )}

      <div className="stats-grid">
        <StatCard title="Total Sales" value={money(salesTotal)} icon="💰" tone="blue" />
        <StatCard title="Gross Profit" value={money(profit)} icon="📈" tone="green" />
        <StatCard title="Expenses" value={money(expensesTotal)} icon="💸" tone="red" />
        <StatCard title="Deposits" value={money(depositsTotal)} icon="🏦" tone="purple" />
      </div>

      <div className="dashboard-grid">
        <div className="panel">
          <div className="panel-header"><div><h3>Sales Summary</h3><span>{data.sales.length} transactions</span></div></div>
          <div className="report-list">
            {data.sales.slice(0, 20).map((x) => (
              <div className="report-row" key={x.id}><div><strong>{x.sale_type}</strong><span>{formatDate(x.sale_date)}</span></div><strong>{money(x.sales_total)}</strong></div>
            ))}
          </div>
        </div>
        <div className="panel">
          <div className="panel-header"><div><h3>Expenses Summary</h3><span>{data.expenses.length} records</span></div></div>
          <div className="report-list">
            {data.expenses.slice(0, 20).map((x) => (
              <div className="report-row" key={x.id}><div><strong>{x.expense_item}</strong><span>{x.category}</span></div><strong>{money(x.amount)}</strong></div>
            ))}
          </div>
        </div>
      </div>

      <div className="stats-grid">
        <StatCard title="Credit Sales" value={money(data.credits.reduce((a,x)=>a+number(x.original_amount),0))} icon="💳" tone="orange" />
        <StatCard title="Outstanding Debt" value={money(creditBalance)} icon="📒" tone="red" />
        <StatCard title="Staff Present" value={data.attendance.filter(a => !!a.check_in_at).length} icon="👥" tone="green" />
        <StatCard title="Attendance Records" value={data.attendance.length} icon="◷" tone="purple" />
      </div>

      <div className="panel">
        <div className="panel-header"><div><h3>Staff Attendance — Daily Report</h3><span>Muda wa kuingia/kutoka unachukuliwa automatically.</span></div></div>
        <div className="table-wrap"><table><thead><tr><th>Staff</th><th>Kuingia</th><th>Kutoka</th><th>Late</th><th>Working Hours</th><th>Status</th></tr></thead><tbody>
        {data.attendance.map(a=>{const cin=a.check_in_at?new Date(a.check_in_at):null;const cout=a.check_out_at?new Date(a.check_out_at):null;const late=cin?Math.max(0,Math.round((cin-(new Date(cin).setHours(8,30,0,0)))/60000)):0;const mins=cin&&cout?Math.max(0,Math.round((cout-cin)/60000)):0;return <tr key={a.id}><td><strong>{a.staff?.staff_name||"-"}</strong></td><td>{cin?cin.toLocaleTimeString("en-TZ",{hour:"2-digit",minute:"2-digit",hour12:false}):"-"}</td><td>{cout?cout.toLocaleTimeString("en-TZ",{hour:"2-digit",minute:"2-digit",hour12:false}):"-"}</td><td>{late} min</td><td>{cin&&cout?`${Math.floor(mins/60)}h ${mins%60}m`:"-"}</td><td>{cout?"PRESENT":"IN WORK"}</td></tr>})}
        {!data.attendance.length&&<tr><td colSpan="6">Hakuna attendance ya siku hii.</td></tr>}
        </tbody></table></div>
      </div>

      <div className="panel">
        <div className="panel-header"><div><h3>Madeni / Credit</h3><span>Wateja waliobaki na salio la deni.</span></div></div>
        <div className="table-wrap"><table><thead><tr><th>Mteja</th><th>Deni</th><th>Amelipa</th><th>Salio</th><th>Status</th></tr></thead><tbody>
        {data.credits.map(c=><tr key={c.id}><td><strong>{c.customer_name}</strong></td><td>{money(c.original_amount)}</td><td>{money(c.paid_amount)}</td><td>{money(c.balance)}</td><td>{c.status}</td></tr>)}
        {!data.credits.length&&<tr><td colSpan="5">Hakuna madeni kwenye kipindi hiki.</td></tr>}
        </tbody></table></div>
      </div>

      <div className="panel summary-panel">
        <h3>Net Position</h3>
        <div className="net-number">{money(salesTotal - expensesTotal)}</div>
        <p>Hii ni Sales minus Expenses kwa kipindi kilichochaguliwa.</p>
      </div>
    </div>
  );
}


function CreditPage({ businessId, staff, refresh }) {
  const [credits, setCredits] = useState([]);
  const [products, setProducts] = useState([]);
  const [services, setServices] = useState([]);
  const [business, setBusiness] = useState(null);
  const [items, setItems] = useState([{ item_type: "PRODUCT", product_id: "", service_id: "", quantity: "1", unit_price: "" }]);
  const [form, setForm] = useState({
    customer_name: "",
    customer_phone: "",
    paid_amount: "0",
    due_date: "",
    note: "",
  });
  const [paying, setPaying] = useState(null);
  const [paymentAmount, setPaymentAmount] = useState("");
  const [search, setSearch] = useState("");
  const [busy, setBusy] = useState(false);
  const [printBill, setPrintBill] = useState(null);

  async function load() {
    if (!businessId) return;
    const [{ data: c, error }, { data: p }, { data: sv }, { data: b }] = await Promise.all([
      supabase.from("credit_transactions").select("*, sales(sale_date, sale_type, quantity, selling_price, products(product_name), services(service_name))").eq("business_id", businessId).order("due_date", { ascending: true }),
      supabase.from("products").select("*").eq("business_id", businessId).eq("active", true).order("product_name"),
      supabase.from("services").select("*").eq("business_id", businessId).eq("active", true).order("service_name"),
      supabase.from("businesses").select("business_name,phone,email,address,logo_url").eq("id", businessId).maybeSingle(),
    ]);
    if (error) { alert(error.message); return; }
    setCredits(c || []); setProducts(p || []); setServices(sv || []); setBusiness(b || null);
  }
  useEffect(() => { load(); }, [businessId, refresh]);

  function getItemLabel(item) {
    if (item.item_type === "SERVICE") return services.find(x => x.id === item.service_id)?.service_name || "Service";
    return products.find(x => x.id === item.product_id)?.product_name || "Product";
  }

  function getItemPrice(item) {
    if (item.item_type === "SERVICE") return number(item.unit_price || services.find(x => x.id === item.service_id)?.selling_price);
    return number(item.unit_price || products.find(x => x.id === item.product_id)?.sell_per_each || products.find(x => x.id === item.product_id)?.selling_price);
  }

  function itemTotal(item) {
    return number(item.quantity) * getItemPrice(item);
  }

  const billTotal = items.reduce((sum, item) => sum + itemTotal(item), 0);
  const paidNow = number(form.paid_amount);
  const billBalance = Math.max(0, billTotal - paidNow);

  function updateItem(index, patch) {
    setItems(prev => prev.map((item, i) => i === index ? { ...item, ...patch } : item));
  }

  function addItemRow() {
    setItems(prev => [...prev, { item_type: "PRODUCT", product_id: "", service_id: "", quantity: "1", unit_price: "" }]);
  }

  function removeItemRow(index) {
    setItems(prev => prev.length === 1 ? prev : prev.filter((_, i) => i !== index));
  }

  function resetBillForm() {
    setForm({ customer_name: "", customer_phone: "", paid_amount: "0", due_date: "", note: "" });
    setItems([{ item_type: "PRODUCT", product_id: "", service_id: "", quantity: "1", unit_price: "" }]);
  }

  function buildStoredNote(lineItems, note) {
    const payload = lineItems.map(x => ({
      type: x.item_type,
      name: getItemLabel(x),
      quantity: number(x.quantity),
      unit_price: getItemPrice(x),
      total: itemTotal(x),
    }));
    return `BILL_ITEMS::${JSON.stringify(payload)}${note ? `\n${note}` : ""}`;
  }

  function parseBillItems(row) {
    const note = String(row?.note || "");
    if (note.startsWith("BILL_ITEMS::")) {
      const firstLine = note.split("\n")[0].replace("BILL_ITEMS::", "");
      try {
        return JSON.parse(firstLine);
      } catch {}
    }
    return [{
      type: row?.product_id ? "PRODUCT" : row?.service_id ? "SERVICE" : "OTHER",
      name: itemName(row),
      quantity: row?.sales?.quantity || 1,
      unit_price: number(row?.original_amount) / Math.max(1, number(row?.sales?.quantity || 1)),
      total: number(row?.original_amount),
    }];
  }

  function printBillWindow(rowOrBill) {
    const row = rowOrBill.row || rowOrBill;
    const lineItems = rowOrBill.items || parseBillItems(row);
    const total = number(row.original_amount ?? lineItems.reduce((a, x) => a + number(x.total), 0));
    const paid = number(row.paid_amount);
    const balance = Math.max(0, total - paid);
    const billNo = String(row.id || Date.now()).slice(-8).toUpperCase();
    const popup = window.open("", "_blank", "width=760,height=900");
    if (!popup) { alert("Browser imezuia popup. Ruhusu popups kwa mfumo huu kisha jaribu tena."); return; }
    const rowsHtml = lineItems.map((x, i) => `<tr><td>${i + 1}</td><td>${escapeHtml(x.name || "-")}</td><td style="text-align:center">${number(x.quantity)}</td><td style="text-align:right">${money(x.unit_price)}</td><td style="text-align:right">${money(x.total)}</td></tr>`).join("");
    const noteText = String(row.note || "").replace(/^BILL_ITEMS::[^\n]*\n?/, "");
    popup.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Bill ${billNo}</title><style>
      *{box-sizing:border-box}body{font-family:Arial,sans-serif;margin:0;padding:28px;color:#172033;background:#fff}.receipt{max-width:700px;margin:auto}.head{text-align:center;border-bottom:2px solid #172033;padding-bottom:16px;margin-bottom:18px}.logo{max-width:72px;max-height:72px;margin-bottom:8px}.head h1{margin:0 0 6px;font-size:24px}.head p{margin:3px 0;font-size:12px;color:#526074}.bill-meta{display:flex;justify-content:space-between;gap:20px;font-size:12px;margin-bottom:18px}.customer{border:1px solid #dce3ed;border-radius:8px;padding:10px;margin-bottom:18px}.customer strong{display:block;margin-bottom:4px}.customer span{font-size:12px;color:#526074}.items{width:100%;border-collapse:collapse;font-size:12px}.items th,.items td{border-bottom:1px solid #e6eaf0;padding:9px 6px}.items th{background:#f5f7fa;text-align:left}.totals{margin-left:auto;width:300px;margin-top:18px;font-size:13px}.totals div{display:flex;justify-content:space-between;padding:5px 0}.totals .balance{border-top:2px solid #172033;margin-top:5px;padding-top:10px;font-size:17px;font-weight:800}.note{margin-top:20px;border:1px dashed #cbd5e1;padding:10px;font-size:12px}.footer{text-align:center;margin-top:28px;font-size:11px;color:#667085}.stamp{display:inline-block;border:2px solid #b42318;color:#b42318;padding:6px 12px;border-radius:6px;font-weight:800;transform:rotate(-3deg);margin-top:14px}@media print{body{padding:0}.receipt{max-width:none}}
    </style></head><body><div class="receipt"><div class="head">${business?.logo_url ? `<img class="logo" src="${escapeHtml(business.logo_url)}" alt="logo">` : ""}<h1>${escapeHtml(business?.business_name || "Bless Stationery")}</h1><p>${escapeHtml(business?.address || "")}</p><p>${escapeHtml(business?.phone || "")} ${business?.email ? `· ${escapeHtml(business.email)}` : ""}</p><h2 style="margin:14px 0 0;font-size:18px">BILL YA MKOPO</h2></div><div class="bill-meta"><div><strong>Bill No:</strong> ${billNo}</div><div><strong>Tarehe:</strong> ${formatDate(new Date().toISOString())}</div><div><strong>Due:</strong> ${escapeHtml(row.due_date || "-")}</div></div><div class="customer"><strong>Mteja: ${escapeHtml(row.customer_name || "-")}</strong><span>${escapeHtml(row.customer_phone || "Hakuna simu")}</span></div><table class="items"><thead><tr><th>#</th><th>Kitu / Huduma</th><th style="text-align:center">Qty</th><th style="text-align:right">Bei</th><th style="text-align:right">Jumla</th></tr></thead><tbody>${rowsHtml}</tbody></table><div class="totals"><div><span>Jumla ya Bill</span><strong>${money(total)}</strong></div><div><span>Amelipa</span><strong>${money(paid)}</strong></div><div class="balance"><span>ANAYODAIWA</span><strong>${money(balance)}</strong></div></div>${noteText ? `<div class="note"><strong>Maelezo:</strong><br>${escapeHtml(noteText).replace(/\n/g,"<br>")}</div>` : ""}<div style="text-align:center"><div class="stamp">${balance > 0 ? "MKOPONI" : "IMELIPWA"}</div></div><div class="footer">Asante kwa kufanya biashara na ${escapeHtml(business?.business_name || "Bless Stationery")}.</div></div><script>window.onload=()=>{window.focus();window.print();}</script></body></html>`);
    popup.document.close();
  }

  async function addDebt(e) {
    e.preventDefault();
    if (!form.customer_name.trim()) { alert("Weka jina la mteja."); return; }
    if (!form.due_date) { alert("Chagua tarehe ya mwisho ya kulipa."); return; }
    const validItems = items.filter(x => (x.item_type === "PRODUCT" ? x.product_id : x.service_id) && number(x.quantity) > 0 && getItemPrice(x) > 0);
    if (!validItems.length) { alert("Ongeza angalau bidhaa/huduma moja yenye kiasi na bei."); return; }
    if (paidNow < 0 || paidNow > billTotal) { alert("Kiasi alicholipa hakiwezi kuzidi jumla ya bill."); return; }
    setBusy(true);
    const balance = Math.max(0, billTotal - paidNow);
    const { data, error } = await supabase.from("credit_transactions").insert({
      business_id: businessId,
      customer_name: form.customer_name.trim(),
      customer_phone: form.customer_phone.trim() || null,
      product_id: null,
      service_id: null,
      original_amount: billTotal,
      paid_amount: paidNow,
      balance,
      status: balance <= 0 ? "PAID" : paidNow > 0 ? "PARTIAL" : "UNPAID",
      due_date: form.due_date,
      note: buildStoredNote(validItems, form.note.trim()),
      created_by: staff?.id || null,
    }).select().single();
    if (error) {
      alert(error.message);
    } else {
      const bill = { ...data, items: validItems.map(x => ({ name: getItemLabel(x), quantity: number(x.quantity), unit_price: getItemPrice(x), total: itemTotal(x) })) };
      resetBillForm();
      await load();
      setPrintBill(bill);
    }
    setBusy(false);
  }

  async function payDebt(row) {
    const pay = number(paymentAmount);
    if (pay <= 0 || pay > number(row.balance)) { alert("Kiasi cha malipo si sahihi."); return; }
    const paid = number(row.paid_amount) + pay;
    const balance = Math.max(0, number(row.original_amount) - paid);
    const { error } = await supabase.from("credit_transactions").update({ paid_amount: paid, balance, status: balance <= 0 ? "PAID" : "PARTIAL", last_payment_at: new Date().toISOString() }).eq("id", row.id).eq("business_id", businessId);
    if (error) { alert(error.message); return; }
    setPaying(null); setPaymentAmount(""); await load();
    if (balance <= 0) sendPaidReceiptWhatsApp({ ...row, paid_amount: paid, balance: 0, status: "PAID" });
  }

  function itemName(x) {
    return x.sales?.products?.product_name || x.sales?.services?.service_name || products.find(p=>p.id===x.product_id)?.product_name || services.find(s=>s.id===x.service_id)?.service_name || "Multiple Items";
  }
  function getWhatsAppPhone(row) {
    let phone = String(row?.customer_phone || "").replace(/[^0-9]/g, "");
    if (phone.startsWith("0")) phone = "255" + phone.slice(1);
    return phone;
  }

  function sendReminder(row) {
    const phone = getWhatsAppPhone(row);
    if (!phone) { alert("Mteja hana namba ya simu."); return; }
    const msg = `Habari ${row.customer_name}, tunakukumbusha kuwa una deni la TZS ${money(row.balance)} katika Bless Stationery. Tarehe ya mwisho ya malipo ni ${row.due_date || "leo"}. Tafadhali lipa kwa wakati. Asante.`;
    window.open(`https://wa.me/${phone}?text=${encodeURIComponent(msg)}`, "_blank");
  }

  function sendPaidReceiptWhatsApp(row) {
    const phone = getWhatsAppPhone(row);
    if (!phone) { alert("Deni limekwisha, lakini mteja hana namba ya WhatsApp iliyohifadhiwa."); return; }
    const lineItems = parseBillItems(row);
    const billNo = String(row.id || Date.now()).slice(-8).toUpperCase();
    const total = number(row.original_amount);
    const paid = number(row.paid_amount);
    const itemText = lineItems.map((x, i) => `${i + 1}. ${x.name || "Item"} × ${number(x.quantity)} = TZS ${money(x.total)}`).join("\n");
    const businessName = business?.business_name || "Bless Stationery";
    const msg = `RECEIPT YA MALIPO YA DENI\n\n${businessName}\nReceipt No: ${billNo}\nMteja: ${row.customer_name || "-"}\nTarehe: ${new Date().toLocaleDateString("en-GB")}\n\nVitu:\n${itemText}\n\nJumla ya deni: TZS ${money(total)}\nJumla iliyolipwa: TZS ${money(paid)}\nSalio: TZS 0\nSTATUS: IMELIPWA KAMILI ✓\n\nAsante kwa kufanya biashara nasi.`;
    window.open(`https://wa.me/${phone}?text=${encodeURIComponent(msg)}`, "_blank");
  }

  const filtered = credits.filter(x => `${x.customer_name} ${x.customer_phone || ""}`.toLowerCase().includes(search.toLowerCase()));
  const totalDebt = filtered.reduce((a,x)=>a+number(x.balance),0);
  const totalOriginal = filtered.reduce((a,x)=>a+number(x.original_amount),0);
  const totalPaid = filtered.reduce((a,x)=>a+number(x.paid_amount),0);
  const overdueCount = filtered.filter(x => number(x.balance) > 0 && x.due_date && new Date(`${x.due_date}T23:59:59`) < new Date()).length;
  function creditStatus(row) { if (number(row.balance) <= 0) return "PAID"; if (row.due_date && new Date(`${row.due_date}T23:59:59`) < new Date()) return "OVERDUE"; return row.status || "UNPAID"; }

  return <div>
    <PageTitle title="Madeni / Credit" subtitle="Tengeneza bill yenye vitu vingi, mpe mteja receipt ya mkopo na fuatilia salio." />
    <div className="stats-grid">
      <StatCard title="Madeni Yote" value={money(totalDebt)} icon="💳" tone="red" />
      <StatCard title="Jumla ya Mikopo" value={money(totalOriginal)} icon="📒" tone="blue" />
      <StatCard title="Yaliyolipwa" value={money(totalPaid)} icon="✓" tone="green" />
      <StatCard title="Wadaiwa" value={filtered.filter(x=>number(x.balance)>0).length} icon="👥" tone="orange" />
      <StatCard title="Overdue" value={overdueCount} icon="⏰" tone="red" />
    </div>

    <div className="panel credit-bill-builder">
      <div className="panel-header"><div><h3>🧾 Tengeneza Bill ya Deni — Multiple Items</h3><span>Mteja anaweza kuchukua bidhaa/huduma nyingi kwenye bill moja.</span></div></div>
      <form onSubmit={addDebt}>
        <div className="form-grid">
          <Field label="Jina la Mteja" value={form.customer_name} onChange={v=>setForm({...form,customer_name:v})} required />
          <Field label="Simu ya Mteja" value={form.customer_phone} onChange={v=>setForm({...form,customer_phone:v})} placeholder="2557XXXXXXXX" />
          <Field label="Deni Litalipwa Tarehe" type="date" value={form.due_date} onChange={v=>setForm({...form,due_date:v})} required />
          <Field label="Kiasi Alicholipa Sasa" type="number" value={form.paid_amount} onChange={v=>setForm({...form,paid_amount:v})} />
        </div>
        <div className="bill-items-header"><strong>Vitu vya Mteja</strong><button type="button" className="secondary-btn small-btn" onClick={addItemRow}>+ Ongeza Kitu</button></div>
        <div className="bill-items-table">
          <div className="bill-item-row bill-item-head"><span>#</span><span>Aina</span><span>Kitu / Huduma</span><span>Qty</span><span>Bei</span><span>Jumla</span><span></span></div>
          {items.map((item, index) => {
            const options = item.item_type === "PRODUCT" ? products.map(p=>({value:p.id,label:p.product_name,price:p.sell_per_each ?? p.selling_price})) : services.map(s=>({value:s.id,label:s.service_name,price:s.selling_price}));
            const selected = options.find(o => o.value === (item.item_type === "PRODUCT" ? item.product_id : item.service_id));
            return <div className="bill-item-row" key={index}>
              <span className="bill-line-no">{index + 1}</span>
              <select value={item.item_type} onChange={e=>updateItem(index,{item_type:e.target.value,product_id:"",service_id:"",unit_price:""})}><option value="PRODUCT">Product</option><option value="SERVICE">Service</option></select>
              <select value={item.item_type === "PRODUCT" ? item.product_id : item.service_id} onChange={e=>{ const patch=item.item_type === "PRODUCT" ? {product_id:e.target.value,service_id:""} : {service_id:e.target.value,product_id:""}; const found=options.find(o=>o.value===e.target.value); updateItem(index,{...patch,unit_price:found?.price ?? ""}); }}><option value="">Chagua...</option>{options.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}</select>
              <input type="number" min="1" step="1" value={item.quantity} onChange={e=>updateItem(index,{quantity:e.target.value})} />
              <input type="number" min="0" step="0.01" value={item.unit_price} onChange={e=>updateItem(index,{unit_price:e.target.value})} placeholder="Bei" />
              <strong className="bill-line-total">{money(itemTotal(item))}</strong>
              <button type="button" className="icon-btn danger-icon" onClick={()=>removeItemRow(index)} disabled={items.length===1}>×</button>
            </div>;
          })}
        </div>
        <div className="bill-summary-box">
          <div><span>Jumla ya Bill</span><strong>{money(billTotal)}</strong></div>
          <div><span>Amelipa Sasa</span><strong>{money(paidNow)}</strong></div>
          <div className="bill-balance"><span>ANAYODAIWA</span><strong>{money(billBalance)}</strong></div>
        </div>
        <div className="field" style={{marginTop:12}}><label>Maelezo / Note</label><textarea rows="2" value={form.note} onChange={e=>setForm({...form,note:e.target.value})} placeholder="Mfano: Vitu vya ofisini, bill ya mwezi..." /></div>
        <div className="form-actions"><button type="button" className="secondary-btn" onClick={resetBillForm}>Clear</button><button className="primary-btn" disabled={busy}>{busy ? "Inahifadhi..." : "✓ Hifadhi Deni + Tengeneza Bill"}</button></div>
      </form>
    </div>

    <div className="panel">
      <div className="toolbar"><input className="search-input" value={search} onChange={e=>setSearch(e.target.value)} placeholder="Tafuta mteja..." /><button className="secondary-btn small-btn" onClick={() => exportToExcel(filtered.map(x=>({Customer:x.customer_name,Phone:x.customer_phone||"",Item:itemName(x),Original:x.original_amount,Paid:x.paid_amount,Balance:x.balance,"Due Date":x.due_date||"",Status:creditStatus(x)})), `Bless-Stationery-Credits-${todayDateInput()}.xlsx`, "Credits")}>⬇ Export</button></div>
      <div className="table-wrap"><table><thead><tr><th>Mteja</th><th>Bill / Vitu</th><th>Deni</th><th>Amelipa</th><th>Salio</th><th>Due Date</th><th>Status</th><th>Action</th></tr></thead><tbody>
        {filtered.map(x=><tr key={x.id}><td><strong>{x.customer_name}</strong><br/><small>{x.customer_phone || "-"}</small></td><td>{itemName(x)}{String(x.note||"").startsWith("BILL_ITEMS::") && <small className="bill-multiple-badge">Multiple Items</small>}</td><td>{money(x.original_amount)}</td><td>{money(x.paid_amount)}</td><td><strong>{money(x.balance)}</strong></td><td>{x.due_date || "-"}</td><td><span className={`status-chip ${creditStatus(x).toLowerCase()}`}>{creditStatus(x)}</span></td><td><div style={{display:"flex",gap:6,flexWrap:"wrap"}}><button className="secondary-btn small-btn" onClick={()=>printBillWindow(x)}>🖨 Print Bill</button>{number(x.balance)>0 ? <>{paying===x.id ? <div className="debt-payment-box"><div className="debt-payment-label">Ingiza kiasi alicholipa</div><input type="number" min="1" max={number(x.balance)} step="0.01" value={paymentAmount} onChange={e=>setPaymentAmount(e.target.value)} placeholder="Mfano 10000" /><div className="debt-payment-balance">Salio baada ya malipo: <strong>{money(Math.max(0, number(x.balance) - number(paymentAmount)))}</strong></div><div className="debt-payment-actions"><button className="primary-btn" onClick={()=>payDebt(x)}>Lipa</button><button className="secondary-btn" onClick={()=>{setPaying(null);setPaymentAmount("")}}>X</button></div></div> : <button className="secondary-btn small-btn" onClick={()=>{setPaying(x.id);setPaymentAmount("")}}>+ Malipo</button>}<button className="secondary-btn small-btn" onClick={()=>sendReminder(x)}>📲 Kumbusha</button></> : <button className="secondary-btn small-btn" onClick={()=>sendPaidReceiptWhatsApp(x)}>📲 Tuma Receipt</button>}</div></td></tr>)}
        {!filtered.length && <tr><td colSpan="8">Hakuna madeni yaliyopatikana.</td></tr>}
      </tbody></table></div>
    </div>

    {printBill && <div className="modal-backdrop" onClick={()=>setPrintBill(null)}><div className="modal-card bill-created-modal" onClick={e=>e.stopPropagation()}><div className="panel-header"><div><h3>✅ Bill imehifadhiwa</h3><span>Mteja sasa ana deni la {money(printBill.balance)}.</span></div><button className="icon-btn" onClick={()=>setPrintBill(null)}>×</button></div><div className="bill-preview-summary"><div><span>Mteja</span><strong>{printBill.customer_name}</strong></div><div><span>Jumla</span><strong>{money(printBill.original_amount)}</strong></div><div><span>Amelipa</span><strong>{money(printBill.paid_amount)}</strong></div><div><span>Anayodaiwa</span><strong>{money(printBill.balance)}</strong></div></div><div className="form-actions" style={{marginTop:18}}><button className="secondary-btn" onClick={()=>setPrintBill(null)}>Funga</button><button className="primary-btn" onClick={()=>printBillWindow(printBill)}>🖨 Print Bill ya Mteja</button></div></div></div>}
  </div>;
}

function AttendancePage({ businessId, staff, refresh }) {
  const [staffRows, setStaffRows] = useState([]);
  const [attendance, setAttendance] = useState([]);
  const [todayRecord, setTodayRecord] = useState(null);
  const [busy, setBusy] = useState(false);
  const [date, setDate] = useState(new Date().toISOString().slice(0,10));
  
  async function load() {
    if (!businessId) return;
    const [{data:s},{data:a,error}] = await Promise.all([
      supabase.from("staff").select("*").eq("business_id",businessId).eq("active",true).order("staff_name"),
      supabase.from("staff_attendance").select("*, staff(staff_name,role)").eq("business_id",businessId).eq("attendance_date",date).order("check_in_at",{ascending:true})
    ]);
    if(error){alert(error.message);return;}
    setStaffRows(s||[]); setAttendance(a||[]);
    setTodayRecord((a||[]).find(x=>x.staff_id===staff?.id)||null);
  }
  useEffect(()=>{load()},[businessId,date,refresh,staff?.id]);

  async function checkIn(){
    if(!staff?.id){alert("Staff profile haijapatikana.");return;}
    if(todayRecord){alert("Tayari ume-sign in leo.");return;}
    setBusy(true);
    const {error}=await supabase.from("staff_attendance").insert({business_id:businessId,staff_id:staff.id,attendance_date:new Date().toLocaleDateString("en-CA",{timeZone:"Africa/Dar_es_Salaam"}),status:"PRESENT"});
    if(error) alert(error.message); else await load(); setBusy(false);
  }
  async function checkOut(){
    if(!todayRecord){alert("Huja-sign in leo.");return;}
    if(todayRecord.check_out_at){alert("Tayari ume-sign out leo.");return;}
    setBusy(true);
    const {error}=await supabase.from("staff_attendance").update({check_out_at:new Date().toISOString()}).eq("id",todayRecord.id).eq("business_id",businessId);
    if(error) alert(error.message); else await load(); setBusy(false);
  }
  function localTime(v){return v?new Date(v).toLocaleTimeString("en-TZ",{hour:"2-digit",minute:"2-digit",hour12:false}):"-"}
  function minutesLate(v){if(!v)return 0;const d=new Date(v);const scheduled=new Date(d);scheduled.setHours(8,30,0,0);return Math.max(0,Math.round((d-scheduled)/60000))}
  function worked(v1,v2){if(!v1||!v2)return "-";return `${Math.floor((new Date(v2)-new Date(v1))/3600000)}h ${Math.floor(((new Date(v2)-new Date(v1))%3600000)/60000)}m`}
  return <div>
    <PageTitle title="Staff Attendance" subtitle="Staff wana-sign in/out; muda unachukuliwa automatically na mfumo." />
    <div className="stats-grid"><StatCard title="Kuingia" value="08:30" icon="🌅" tone="blue" /><StatCard title="Kutoka" value="21:30" icon="🌙" tone="purple" /><StatCard title="Present" value={attendance.filter(a=>!!a.check_in_at).length} icon="✓" tone="green" /><StatCard title="Absent" value={Math.max(0, staffRows.filter(s=>!attendance.some(a=>a.staff_id===s.id)).length)} icon="◷" tone="red" /></div>
    {staff?.active && <div className="panel"><div className="panel-header"><div><h3>{staff.staff_name} — Leo</h3><span>System time ndiyo unaotumika; staff haandiki muda.</span></div><div style={{display:"flex",gap:8}}><button className="primary-btn" disabled={busy||!!todayRecord} onClick={checkIn}>✓ SIGN IN</button><button className="secondary-btn" disabled={busy||!todayRecord||!!todayRecord.check_out_at} onClick={checkOut}>↪ SIGN OUT</button></div></div>{todayRecord&&<div className="notice">Kuingia: <strong>{localTime(todayRecord.check_in_at)}</strong> · Kutoka: <strong>{localTime(todayRecord.check_out_at)}</strong> · Late: <strong>{minutesLate(todayRecord.check_in_at)} min</strong></div>}</div>}
    <div className="panel"><div className="toolbar"><input type="date" value={date} onChange={e=>setDate(e.target.value)} /></div><div className="table-wrap"><table><thead><tr><th>Staff</th><th>Kuingia</th><th>Kutoka</th><th>Late</th><th>Working Hours</th><th>Status</th></tr></thead><tbody>{staffRows.map(s=>{const a=attendance.find(x=>x.staff_id===s.id);return <tr key={s.id}><td><strong>{s.staff_name}</strong></td><td>{localTime(a?.check_in_at)}</td><td>{localTime(a?.check_out_at)}</td><td>{a?`${minutesLate(a.check_in_at)} min`:"-"}</td><td>{worked(a?.check_in_at,a?.check_out_at)}</td><td>{a?(a.check_out_at?(minutesLate(a.check_in_at)>0?"LATE":"PRESENT"):"IN WORK"):"ABSENT"}</td></tr>})}</tbody></table></div></div>
  </div>;
}

function StaffPage({ businessId, refresh }) {
  const [staff, setStaff] = useState([]);
  const [showForm, setShowForm] = useState(false);
  const [search, setSearch] = useState("");
  const [busy, setBusy] = useState(false);
  const [form, setForm] = useState({
    staff_name: "",
    role: "STAFF",
    active: true,
  });

  useEffect(() => {
    load();
  }, [businessId]);

  async function load() {
    if (!businessId) return;
    const { data, error } = await supabase
      .from("staff")
      .select("*")
      .eq("business_id", businessId)
      .order("created_at", { ascending: true });

    if (error) {
      alert(error.message);
      return;
    }

    setStaff(data || []);
  }

  async function addStaffProfile(e) {
    e.preventDefault();

    if (!form.staff_name.trim()) {
      alert("Weka jina la staff.");
      return;
    }

    setBusy(true);

    const { error } = await supabase.from("staff").insert({
      business_id: businessId,
      staff_name: form.staff_name.trim(),
      role: form.role,
      active: form.active,
    });

    if (error) {
      alert(error.message);
    } else {
      setForm({ staff_name: "", role: "STAFF", active: true });
      setShowForm(false);
      await load();
    }

    setBusy(false);
  }

  async function toggleActive(row) {
    const { error } = await supabase
      .from("staff")
      .update({ active: !row.active })
      .eq("id", row.id)
      .eq("business_id", businessId);

    if (error) alert(error.message);
    else await load();
  }

  async function deleteStaff(row) {
    if (row.role === "ADMIN") {
      alert("Admin account haiwezi kufutwa hapa.");
      return;
    }

    if (!confirm(`Futa staff profile ya ${row.staff_name}?`)) return;

    const { error } = await supabase
      .from("staff")
      .delete()
      .eq("id", row.id)
      .eq("business_id", businessId);

    if (error) alert(error.message);
    else await load();
  }

  const filtered = staff.filter((x) =>
    `${x.staff_name} ${x.role}`.toLowerCase().includes(search.toLowerCase())
  );

  const activeCount = staff.filter((x) => x.active).length;
  const adminCount = staff.filter((x) => x.role === "ADMIN").length;

  return (
    <div>
      <div className="page-title staff-page-title">
        <div>
          <div className="eyebrow">USER MANAGEMENT</div>
          <h1>Staff & Users</h1>
          <p>Simamia profiles za staff na hali ya akaunti zao.</p>
        </div>
        <button className="primary-btn" onClick={() => setShowForm(!showForm)}>
          {showForm ? "Close Form" : "+ Add Staff"}
        </button>
      </div>

      <div className="stats-grid staff-stats">
        <StatCard title="Total Users" value={staff.length} icon="👥" tone="blue" />
        <StatCard title="Active" value={activeCount} icon="✓" tone="green" />
        <StatCard title="Admin" value={adminCount} icon="🛡" tone="purple" />
        <StatCard title="Staff" value={staff.filter((x) => x.role === "STAFF").length} icon="◉" tone="orange" />
      </div>

      {showForm && (
        <div className="panel staff-form-panel">
          <div className="panel-header">
            <div>
              <h3>Add Staff Profile</h3>
              <span>Ongeza taarifa ya staff kwenye mfumo.</span>
            </div>
          </div>

          <form onSubmit={addStaffProfile} className="form-grid">
            <Field
              label="Staff Name"
              value={form.staff_name}
              onChange={(v) => setForm({ ...form, staff_name: v })}
              placeholder="Mfano: John"
              required
            />

            <SelectField
              label="Role"
              value={form.role}
              onChange={(v) => setForm({ ...form, role: v })}
              options={[
                { value: "STAFF", label: "STAFF" },
                { value: "ADMIN", label: "ADMIN" },
              ]}
            />

            <SelectField
              label="Status"
              value={form.active ? "ACTIVE" : "INACTIVE"}
              onChange={(v) => setForm({ ...form, active: v === "ACTIVE" })}
              options={[
                { value: "ACTIVE", label: "ACTIVE" },
                { value: "INACTIVE", label: "INACTIVE" },
              ]}
            />

            <div className="notice full">
              🔐 Hii form inatengeneza <strong>staff profile</strong>. Login email/password ya Supabase Auth inapaswa kuundwa kupitia Authentication; hatuweki secret key kwenye browser.
            </div>

            <div className="form-actions full">
              <button className="primary-btn" disabled={busy}>
                {busy ? "Inahifadhi..." : "Save Staff Profile"}
              </button>
              <button
                type="button"
                className="secondary-btn"
                onClick={() => setShowForm(false)}
              >
                Cancel
              </button>
            </div>
          </form>
        </div>
      )}

      <div className="panel">
        <div className="toolbar staff-toolbar">
          <div>
            <div className="panel-kicker">TEAM DIRECTORY</div>
            <h3 className="toolbar-title">Users & Staff</h3>
          </div>
          <div className="search-box staff-search">
            <span>⌕</span>
            <input
              placeholder="Search staff..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
        </div>

        <div className="staff-directory">
          {filtered.map((row) => (
            <div className="staff-directory-card" key={row.id}>
              <div className="staff-main">
                <div className={`staff-avatar ${row.role === "ADMIN" ? "admin-avatar" : ""}`}>
                  {(row.staff_name || "U").charAt(0).toUpperCase()}
                </div>
                <div>
                  <h3>{row.staff_name}</h3>
                  <div className="staff-meta">
                    <span className={`role-chip ${row.role === "ADMIN" ? "admin" : "staff"}`}>
                      {row.role}
                    </span>
                    <span className={`status-chip ${row.active ? "active" : "inactive"}`}>
                      {row.active ? "Active" : "Inactive"}
                    </span>
                  </div>
                </div>
              </div>

              <div className="staff-id-box">
                <span>Profile ID</span>
                <strong>{row.id.slice(0, 8).toUpperCase()}</strong>
              </div>

              <div className="staff-actions">
                <button
                  className={row.active ? "status-btn deactivate" : "status-btn activate"}
                  onClick={() => toggleActive(row)}
                >
                  {row.active ? "Deactivate" : "Activate"}
                </button>
                {row.role !== "ADMIN" && (
                  <button className="danger-small" onClick={() => deleteStaff(row)}>
                    Delete
                  </button>
                )}
              </div>
            </div>
          ))}

          {!filtered.length && <EmptyState text="Hakuna staff anayefanana na search yako." />}
        </div>
      </div>
    </div>
  );
}

function BusinessSettingsPage({ businessId, business, user, theme, setTheme, density, setDensity, onChanged }) {
  const [form, setForm] = useState({
    business_name: business?.business_name || "",
    phone: business?.phone || "",
    email: business?.email || "",
    address: business?.address || "",
    logo_url: business?.logo_url || "",
  });
  const [busy, setBusy] = useState(false);
  const [passwordBusy, setPasswordBusy] = useState(false);
  const [passwordMessage, setPasswordMessage] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [currentPassword, setCurrentPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirmNewPassword, setConfirmNewPassword] = useState("");
  const [showCurrent, setShowCurrent] = useState(false);
  const [showNew, setShowNew] = useState(false);
  const [showConfirm, setShowConfirm] = useState(false);

  useEffect(() => {
    setForm({
      business_name: business?.business_name || "",
      phone: business?.phone || "",
      email: business?.email || "",
      address: business?.address || "",
      logo_url: business?.logo_url || "",
    });
  }, [business]);

  async function save(e) {
    e.preventDefault();
    if (!form.business_name.trim()) {
      alert("Weka jina la biashara.");
      return;
    }
    setBusy(true);
    const { error } = await supabase
      .from("businesses")
      .update({
        business_name: form.business_name.trim(),
        phone: form.phone.trim() || null,
        email: form.email.trim() || null,
        address: form.address.trim() || null,
        logo_url: form.logo_url.trim() || null,
      })
      .eq("id", businessId);

    if (error) alert(error.message);
    else {
      alert("Business profile imehifadhiwa.");
      onChanged();
    }
    setBusy(false);
  }

  function chooseTheme(value) {
    setTheme(value);
  }

  async function changePassword(e) {
    e.preventDefault();
    setPasswordError("");
    setPasswordMessage("");

    if (!currentPassword) {
      setPasswordError("Weka password yako ya sasa.");
      return;
    }
    if (newPassword.length < 6) {
      setPasswordError("Password mpya lazima iwe na angalau characters 6.");
      return;
    }
    if (newPassword !== confirmNewPassword) {
      setPasswordError("Password mpya hazifanani.");
      return;
    }
    if (newPassword === currentPassword) {
      setPasswordError("Password mpya iwe tofauti na ya sasa.");
      return;
    }

    setPasswordBusy(true);
    try {
      const { error: verifyError } = await supabase.auth.signInWithPassword({
        email: user?.email || "",
        password: currentPassword,
      });

      if (verifyError) {
        setPasswordError("Password ya sasa sio sahihi.");
        return;
      }

      const { error: updateError } = await supabase.auth.updateUser({
        password: newPassword,
      });

      if (updateError) throw updateError;

      setCurrentPassword("");
      setNewPassword("");
      setConfirmNewPassword("");
      setPasswordMessage("Password imebadilishwa kikamilifu.");
    } catch (err) {
      console.error("Password change failed:", err);
      setPasswordError(err?.message || "Imeshindikana kubadilisha password.");
    } finally {
      setPasswordBusy(false);
    }
  }

  const themeOptions = [
    { id: "ocean", name: "Ocean Blue", desc: "Blue + Teal", color: "#0f3d5e", accent: "#0ea5a4" },
    { id: "emerald", name: "Emerald", desc: "Green + Teal", color: "#064e3b", accent: "#10b981" },
    { id: "purple", name: "Royal Purple", desc: "Purple + Violet", color: "#3b1f6f", accent: "#8b5cf6" },
    { id: "slate", name: "Slate", desc: "Dark Slate + Cyan", color: "#172033", accent: "#06b6d4" },
    { id: "rose", name: "Rose", desc: "Wine + Rose", color: "#5f172b", accent: "#e11d48" },
  ];

  return (
    <div>
      <PageTitle title="Settings" subtitle="Dhibiti taarifa za biashara, mwonekano wa mfumo na usalama wa account yako." />

      <div className="settings-grid">
        <div className="panel">
          <div className="panel-header">
            <div>
              <h3>Business Profile</h3>
              <span>Taarifa hizi zinatumika kwenye mfumo na receipt.</span>
            </div>
          </div>
          <form onSubmit={save} className="form-grid">
            <Field label="Business Name" value={form.business_name} onChange={v => setForm({...form,business_name:v})} required />
            <Field label="Phone" value={form.phone} onChange={v => setForm({...form,phone:v})} />
            <Field label="Email" type="email" value={form.email} onChange={v => setForm({...form,email:v})} />
            <Field label="Address" value={form.address} onChange={v => setForm({...form,address:v})} />
            <Field label="Logo URL" value={form.logo_url} onChange={v => setForm({...form,logo_url:v})} placeholder="Optional" />
            <div className="notice full">💡 Jina, simu, address na logo vinaweza kuonekana kwenye receipt kulingana na layout ya receipt.</div>
            <div className="form-actions full"><button className="primary-btn" disabled={busy}>{busy ? "Inahifadhi..." : "Save Business Profile"}</button></div>
          </form>
        </div>

        <div className="panel settings-preview">
          <div className="panel-kicker">RECEIPT PREVIEW</div>
          <div className="settings-logo">{form.logo_url ? <img src={form.logo_url} alt="" /> : <div className="receipt-logo">B</div>}</div>
          <h2>{form.business_name || "Bless Stationery"}</h2>
          <p>{form.address || "Business address"}</p>
          <p>{form.phone || "Phone number"}</p>
          <p>{form.email || "Email"}</p>
        </div>
      </div>

      <div className="settings-section-title">
        <div>
          <h2>🎨 Mwonekano wa Mfumo</h2>
          <p>Chagua rangi ambayo biashara yako itatumia kwenye dashboard, menus na buttons.</p>
        </div>
      </div>

      <div className="theme-grid">
        {themeOptions.map((item) => (
          <button
            type="button"
            key={item.id}
            className={`theme-card ${theme === item.id ? "selected" : ""}`}
            onClick={() => chooseTheme(item.id)}
          >
            <div className="theme-preview" style={{ "--preview-main": item.color, "--preview-accent": item.accent }}>
              <span></span><span></span><span></span>
            </div>
            <div className="theme-card-copy">
              <strong>{item.name}</strong>
              <small>{item.desc}</small>
            </div>
            <div className="theme-check">{theme === item.id ? "✓" : ""}</div>
          </button>
        ))}
      </div>

      <div className="settings-grid settings-grid-bottom">
        <div className="panel">
          <div className="panel-header">
            <div>
              <h3>🖥 Display Preferences</h3>
              <span>Mipangilio hii inahifadhiwa kwenye browser ya kifaa hiki.</span>
            </div>
          </div>
          <div className="preference-row">
            <div>
              <strong>Density ya mfumo</strong>
              <p>Chagua kama tables na cards ziwe compact au ziwe na nafasi zaidi.</p>
            </div>
            <div className="segmented-control">
              <button type="button" className={density === "comfortable" ? "active" : ""} onClick={() => setDensity("comfortable")}>Comfortable</button>
              <button type="button" className={density === "compact" ? "active" : ""} onClick={() => setDensity("compact")}>Compact</button>
            </div>
          </div>
          <div className="notice">ℹ️ Theme na density ni za account hii kwenye browser hii; database yako haihitaji kubadilishwa.</div>
        </div>

        <div className="panel">
          <div className="panel-header">
            <div>
              <h3>🔐 Account Security</h3>
              <span>Email ya account: <strong>{user?.email || "-"}</strong></span>
            </div>
          </div>
          {passwordError && <div className="error-box">{passwordError}</div>}
          {passwordMessage && <div className="success-box">{passwordMessage}</div>}
          <form onSubmit={changePassword} className="password-settings-form">
            <PasswordField label="Password ya sasa" value={currentPassword} onChange={setCurrentPassword} show={showCurrent} setShow={setShowCurrent} />
            <PasswordField label="Password mpya" value={newPassword} onChange={setNewPassword} show={showNew} setShow={setShowNew} />
            <PasswordField label="Confirm password mpya" value={confirmNewPassword} onChange={setConfirmNewPassword} show={showConfirm} setShow={setShowConfirm} />
            <div className="password-hint">Password lazima iwe na angalau characters 6. Tunathibitisha password yako ya sasa kabla ya kuweka mpya.</div>
            <button className="primary-btn" disabled={passwordBusy} type="submit">{passwordBusy ? "Inabadilisha..." : "Badilisha Password"}</button>
          </form>
        </div>
      </div>

      <div className="panel settings-help-panel">
        <div className="panel-header">
          <div>
            <h3>⚙️ Mipangilio mingine muhimu</h3>
            <span>Sehemu ambazo tunaweza kuongeza baadaye bila kubadilisha mfumo wa msingi.</span>
          </div>
        </div>
        <div className="settings-help-grid">
          <div><strong>Receipt Settings</strong><span>80mm / A4, footer message, logo na receipt notes.</span></div>
          <div><strong>Business Defaults</strong><span>Default payment method, tax/VAT display na invoice numbering.</span></div>
          <div><strong>Security</strong><span>Session control, logout other devices na role permissions.</span></div>
          <div><strong>Notifications</strong><span>Low-stock threshold, reminders na in-app alerts.</span></div>
        </div>
      </div>
    </div>
  );
}

function PasswordField({ label, value, onChange, show, setShow }) {
  return (
    <div className="field password-settings-field">
      <label>{label}</label>
      <div className="password-wrap">
        <input type={show ? "text" : "password"} value={value} onChange={(e) => onChange(e.target.value)} required />
        <button type="button" className="show-password" onClick={() => setShow(!show)}>{show ? "Ficha" : "Onyesha"}</button>
      </div>
    </div>
  );
}


function CustomersPage({ businessId, refresh }) {
  const [customers, setCustomers] = useState([]);
  const [credits, setCredits] = useState([]);
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => { load(); }, [businessId, refresh]);

  async function load() {
    if (!businessId) return;
    setLoading(true);
    const [{ data: creditRows, error: creditError }, { data: salesRows }] = await Promise.all([
      supabase.from("credit_transactions").select("*").eq("business_id", businessId).order("created_at", { ascending: false }).limit(2000),
      supabase.from("sales").select("id,sale_date,sales_total,profit,payment_method,quantity,product_id,service_id").eq("business_id", businessId).order("sale_date", { ascending: false }).limit(5000),
    ]);
    if (creditError) {
      alert(creditError.message);
      setLoading(false);
      return;
    }
    const rows = creditRows || [];
    const map = {};
    rows.forEach((r) => {
      const key = `${String(r.customer_name || "").trim().toLowerCase()}|${normalizePhone(r.customer_phone)}`;
      if (!key || key === "|") return;
      if (!map[key]) {
        map[key] = {
          key,
          name: r.customer_name || "-",
          phone: r.customer_phone || "",
          totalDebt: 0,
          balance: 0,
          paid: 0,
          transactions: 0,
          lastDate: r.created_at || r.due_date || null,
        };
      }
      map[key].totalDebt += number(r.original_amount);
      map[key].balance += number(r.balance);
      map[key].paid += number(r.paid_amount);
      map[key].transactions += 1;
      if (new Date(r.created_at || 0) > new Date(map[key].lastDate || 0)) map[key].lastDate = r.created_at;
    });
    const list = Object.values(map).sort((a, b) => b.balance - a.balance || a.name.localeCompare(b.name));
    setCustomers(list);
    setCredits(rows);
    setSelected((current) => current ? list.find((x) => x.key === current.key) || null : null);
    setLoading(false);
  }

  const filtered = customers.filter((c) => {
    const q = search.trim().toLowerCase();
    if (!q) return true;
    return c.name.toLowerCase().includes(q) || String(c.phone).toLowerCase().includes(q);
  });

  const selectedRows = selected
    ? credits.filter((r) => {
        const a = String(r.customer_name || "").trim().toLowerCase();
        const b = String(selected.name || "").trim().toLowerCase();
        return a === b && normalizePhone(r.customer_phone) === normalizePhone(selected.phone);
      })
    : [];

  function printStatement() {
    if (!selected) return;
    const win = window.open("", "_blank", "width=900,height=800");
    if (!win) { alert("Ruhusu pop-ups kisha ujaribu tena."); return; }
    const rows = selectedRows.map((r) => `
      <tr><td>${escapeHtml(formatDate(r.created_at))}</td><td>${escapeHtml(r.note || "Credit")}</td><td class="num">${escapeHtml(money(r.original_amount))}</td><td class="num">${escapeHtml(money(r.paid_amount))}</td><td class="num">${escapeHtml(money(r.balance))}</td><td>${escapeHtml(r.status || "-")}</td></tr>
    `).join("");
    win.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Customer Statement</title><style>body{font-family:Arial;padding:28px;color:#111}h1{margin:0 0 5px}.muted{color:#666}table{width:100%;border-collapse:collapse;margin-top:22px}th,td{border:1px solid #ccc;padding:8px}th{background:#f3f4f6;text-align:left}.num{text-align:right}.summary{margin-top:18px;display:flex;gap:28px}.summary strong{display:block;font-size:18px}@media print{body{padding:0}}</style></head><body><h1>${escapeHtml(selected.name)}</h1><div class="muted">${escapeHtml(selected.phone || "No phone")}</div><div class="summary"><div>Total Credit<strong>${escapeHtml(money(selected.totalDebt))}</strong></div><div>Paid<strong>${escapeHtml(money(selected.paid))}</strong></div><div>Balance<strong>${escapeHtml(money(selected.balance))}</strong></div></div><table><thead><tr><th>Date</th><th>Note</th><th>Debt</th><th>Paid</th><th>Balance</th><th>Status</th></tr></thead><tbody>${rows}</tbody></table><script>window.onload=()=>window.print()</script></body></html>`);
    win.document.close();
  }

  if (loading) return <PageLoading />;

  return (
    <div>
      <PageTitle title="Customers" subtitle="Wateja, historia ya madeni na salio lao kwa sehemu moja." />
      <div className="stats-grid small-stats">
        <StatCard title="Customers" value={customers.length} icon="👥" tone="blue" />
        <StatCard title="Customers Wenye Deni" value={customers.filter((x) => x.balance > 0).length} icon="💳" tone="red" />
        <StatCard title="Total Credit" value={money(customers.reduce((a, x) => a + x.totalDebt, 0))} icon="🧾" tone="purple" />
        <StatCard title="Outstanding" value={money(customers.reduce((a, x) => a + x.balance, 0))} icon="⚠" tone="orange" />
      </div>

      <div className="panel">
        <div className="panel-header">
          <div><h3>Customer Directory</h3><span>Search kwa jina au simu.</span></div>
          <button className="secondary-btn" onClick={load}>↻ Refresh</button>
        </div>
        <input className="customer-search" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Tafuta mteja..." />
        <div className="table-wrap">
          <table>
            <thead><tr><th>Mteja</th><th>Simu</th><th>Transactions</th><th>Total Credit</th><th>Paid</th><th>Balance</th><th>Last Activity</th><th>Action</th></tr></thead>
            <tbody>
              {filtered.map((c) => (
                <tr key={c.key}>
                  <td><strong>{c.name}</strong></td><td>{c.phone || "-"}</td><td>{c.transactions}</td><td>{money(c.totalDebt)}</td><td>{money(c.paid)}</td><td><strong className={c.balance > 0 ? "stock-danger" : "stock-ok"}>{money(c.balance)}</strong></td><td>{formatDate(c.lastDate)}</td>
                  <td><button className="secondary-btn small-btn" onClick={() => setSelected(c)}>View</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {!filtered.length && <EmptyState text="Hakuna customer anayefanana na search yako." />}
      </div>

      {selected && (
        <div className="panel customer-detail-panel">
          <div className="panel-header">
            <div><h3>{selected.name}</h3><span>{selected.phone || "No phone"}</span></div>
            <div className="panel-header-actions"><button className="secondary-btn" onClick={printStatement}>🖨 Print Statement</button><button className="text-btn" onClick={() => setSelected(null)}>Funga</button></div>
          </div>
          <div className="stats-grid small-stats">
            <StatCard title="Total Credit" value={money(selected.totalDebt)} icon="🧾" />
            <StatCard title="Paid" value={money(selected.paid)} icon="✓" tone="green" />
            <StatCard title="Outstanding" value={money(selected.balance)} icon="💳" tone="red" />
            <StatCard title="Transactions" value={selected.transactions} icon="↕" tone="purple" />
          </div>
          <div className="table-wrap">
            <table><thead><tr><th>Date</th><th>Note</th><th>Debt</th><th>Paid</th><th>Balance</th><th>Due Date</th><th>Status</th></tr></thead>
              <tbody>{selectedRows.map((r) => <tr key={r.id}><td>{formatDate(r.created_at)}</td><td>{r.note || "Credit"}</td><td>{money(r.original_amount)}</td><td>{money(r.paid_amount)}</td><td><strong>{money(r.balance)}</strong></td><td>{r.due_date || "-"}</td><td>{r.status || "-"}</td></tr>)}</tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function AnalyticsPage({ businessId, refresh }) {
  const [from, setFrom] = useState(todayDateInput());
  const [to, setTo] = useState(todayDateInput());
  const [data, setData] = useState({ sales: [], expenses: [], deposits: [], credits: [] });
  const [loading, setLoading] = useState(true);

  useEffect(() => { load(); }, [businessId, refresh]);

  async function load() {
    if (!businessId) return;
    setLoading(true);
    const [salesRes, expenseRes, depositRes, creditRes] = await Promise.all([
      supabase.from("sales").select("*").eq("business_id", businessId).order("sale_date", { ascending: false }).limit(5000),
      supabase.from("expenses").select("*").eq("business_id", businessId).order("expense_date", { ascending: false }).limit(5000),
      supabase.from("deposits").select("*").eq("business_id", businessId).order("deposit_date", { ascending: false }).limit(5000),
      supabase.from("credit_transactions").select("*").eq("business_id", businessId).order("created_at", { ascending: false }).limit(5000),
    ]);
    if (salesRes.error) alert(salesRes.error.message);
    if (expenseRes.error) alert(expenseRes.error.message);
    if (depositRes.error) alert(depositRes.error.message);
    if (creditRes.error) alert(creditRes.error.message);
    setData({ sales: salesRes.data || [], expenses: expenseRes.data || [], deposits: depositRes.data || [], credits: creditRes.data || [] });
    setLoading(false);
  }

  const start = new Date(`${from}T00:00:00`);
  const end = new Date(`${to}T23:59:59.999`);
  const inRange = (value) => { const d = new Date(value); return d >= start && d <= end; };
  const sales = data.sales.filter((x) => inRange(x.sale_date));
  const expenses = data.expenses.filter((x) => inRange(x.expense_date));
  const deposits = data.deposits.filter((x) => inRange(x.deposit_date));
  const salesTotal = sales.reduce((a, x) => a + number(x.sales_total), 0);
  const grossProfit = sales.reduce((a, x) => a + number(x.profit), 0);
  const expenseTotal = expenses.reduce((a, x) => a + number(x.amount), 0);
  const netProfit = grossProfit - expenseTotal;
  const creditSales = sales.filter((x) => x.payment_method === "CREDIT").reduce((a, x) => a + number(x.sales_total), 0);
  const cashSales = sales.filter((x) => x.payment_method === "CASH").reduce((a, x) => a + number(x.sales_total), 0);
  const outstanding = data.credits.reduce((a, x) => a + number(x.balance), 0);
  const depositTotal = deposits.reduce((a, x) => a + number(x.amount), 0);
  const maxDay = Math.max(1, ...Array.from({ length: 7 }, (_, i) => {
    const d = new Date(); d.setHours(0,0,0,0); d.setDate(d.getDate() - (6 - i));
    const next = new Date(d); next.setDate(next.getDate() + 1);
    return data.sales.filter((x) => new Date(x.sale_date) >= d && new Date(x.sale_date) < next).reduce((a,x)=>a+number(x.sales_total),0);
  }));
  const paymentMap = {};
  sales.forEach((x) => { paymentMap[x.payment_method || "OTHER"] = (paymentMap[x.payment_method || "OTHER"] || 0) + number(x.sales_total); });

  function printReport() {
    const win = window.open("", "_blank", "width=1000,height=800");
    if (!win) { alert("Ruhusu pop-ups kisha ujaribu tena."); return; }
    win.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Business Profit & Loss</title><style>body{font-family:Arial;padding:28px;color:#111}h1{margin:0 0 5px}.muted{color:#666}.cards{display:flex;gap:12px;margin:18px 0}.card{border:1px solid #ddd;padding:12px;flex:1}.card strong{display:block;font-size:18px;margin-top:5px}table{width:100%;border-collapse:collapse;margin-top:18px}th,td{border:1px solid #ccc;padding:8px;text-align:left}.num{text-align:right}th{background:#f3f4f6}@media print{body{padding:0}}</style></head><body><h1>Bless Stationery - Profit & Loss</h1><div class="muted">Period: ${escapeHtml(from)} to ${escapeHtml(to)}</div><div class="cards"><div class="card">Sales<strong>${escapeHtml(money(salesTotal))}</strong></div><div class="card">Gross Profit<strong>${escapeHtml(money(grossProfit))}</strong></div><div class="card">Expenses<strong>${escapeHtml(money(expenseTotal))}</strong></div><div class="card">Net Profit<strong>${escapeHtml(money(netProfit))}</strong></div></div><table><thead><tr><th>Metric</th><th class="num">Amount</th></tr></thead><tbody><tr><td>Cash Sales</td><td class="num">${escapeHtml(money(cashSales))}</td></tr><tr><td>Credit Sales</td><td class="num">${escapeHtml(money(creditSales))}</td></tr><tr><td>Deposits</td><td class="num">${escapeHtml(money(depositTotal))}</td></tr><tr><td>Outstanding Credit</td><td class="num">${escapeHtml(money(outstanding))}</td></tr></tbody></table><script>window.onload=()=>window.print()</script></body></html>`);
    win.document.close();
  }

  if (loading) return <PageLoading />;
  return (
    <div>
      <PageTitle title="Business Analytics" subtitle="Profit & Loss, payment mix na performance ya biashara." />
      <div className="panel analytics-filter">
        <div className="form-grid">
          <Field label="From" type="date" value={from} onChange={setFrom} />
          <Field label="To" type="date" value={to} onChange={setTo} />
          <div className="form-actions"><button className="secondary-btn" onClick={load}>↻ Refresh</button><button className="primary-btn" onClick={printReport}>🖨 Print Report</button></div>
        </div>
      </div>
      <div className="stats-grid">
        <StatCard title="Sales" value={money(salesTotal)} icon="🛒" tone="blue" />
        <StatCard title="Gross Profit" value={money(grossProfit)} icon="📈" tone="green" />
        <StatCard title="Expenses" value={money(expenseTotal)} icon="💸" tone="red" />
        <StatCard title="Net Profit" value={money(netProfit)} icon="💎" tone={netProfit >= 0 ? "green" : "red"} />
      </div>
      <div className="stats-grid small-stats">
        <StatCard title="Cash Sales" value={money(cashSales)} icon="💵" />
        <StatCard title="Credit Sales" value={money(creditSales)} icon="💳" tone="purple" />
        <StatCard title="Deposits" value={money(depositTotal)} icon="🏦" tone="blue" />
        <StatCard title="Outstanding Debt" value={money(outstanding)} icon="⚠" tone="orange" />
      </div>
      <div className="dashboard-grid">
        <div className="panel">
          <div className="panel-header"><div><h3>Sales - Last 7 Days</h3><span>Muonekano wa haraka wa sales.</span></div></div>
          <div className="analytics-bars">
            {Array.from({ length: 7 }, (_, i) => {
              const d = new Date(); d.setHours(0,0,0,0); d.setDate(d.getDate() - (6 - i));
              const next = new Date(d); next.setDate(next.getDate()+1);
              const total = data.sales.filter((x) => new Date(x.sale_date) >= d && new Date(x.sale_date) < next).reduce((a,x)=>a+number(x.sales_total),0);
              return <div className="analytics-bar-row" key={d.toISOString()}><span>{d.toLocaleDateString("en-TZ", { weekday: "short" })}</span><div className="analytics-bar-track"><div className="analytics-bar-fill" style={{ width: `${Math.min(100, total / maxDay * 100)}%` }} /></div><strong>{money(total)}</strong></div>;
            })}
          </div>
        </div>
        <div className="panel">
          <div className="panel-header"><div><h3>Payment Mix</h3><span>Sales kwa payment method.</span></div></div>
          <div className="mini-table">
            {Object.entries(paymentMap).sort((a,b)=>b[1]-a[1]).map(([method,total]) => <div className="mini-row" key={method}><div><strong>{method}</strong><span>{sales.length} sales range</span></div><strong>{money(total)}</strong></div>)}
          </div>
          {!Object.keys(paymentMap).length && <EmptyState text="Hakuna sales kwenye kipindi hiki." />}
        </div>
      </div>
      <div className="panel">
        <div className="panel-header"><div><h3>Expense Breakdown</h3><span>{expenses.length} expense records</span></div></div>
        <div className="mini-table">
          {Object.entries(expenses.reduce((acc, x) => { const k = x.category || "Other"; acc[k] = (acc[k] || 0) + number(x.amount); return acc; }, {})).sort((a,b)=>b[1]-a[1]).map(([cat,total]) => <div className="mini-row" key={cat}><div><strong>{cat}</strong><span>Expense category</span></div><strong>{money(total)}</strong></div>)}
        </div>
      </div>
    </div>
  );
}

function CashClosingPage({ businessId, refresh }) {
  const [date, setDate] = useState(todayDateInput());
  const [actualCash, setActualCash] = useState("");
  const [records, setRecords] = useState({ sales: [], expenses: [], deposits: [], credits: [] });
  const [loading, setLoading] = useState(true);

  useEffect(() => { load(); }, [businessId, refresh, date]);
  useEffect(() => {
    try {
      const saved = localStorage.getItem(`bless_cash_count_${businessId}_${date}`);
      setActualCash(saved || "");
    } catch {}
  }, [businessId, date]);

  async function load() {
    if (!businessId) return;
    setLoading(true);
    const start = `${date}T00:00:00`;
    const end = `${date}T23:59:59.999`;
    const [salesRes, expenseRes, depositRes, creditRes] = await Promise.all([
      supabase.from("sales").select("*").eq("business_id", businessId).gte("sale_date", start).lte("sale_date", end),
      supabase.from("expenses").select("*").eq("business_id", businessId).gte("expense_date", start).lte("expense_date", end),
      supabase.from("deposits").select("*").eq("business_id", businessId).gte("deposit_date", start).lte("deposit_date", end),
      supabase.from("credit_transactions").select("*").eq("business_id", businessId),
    ]);
    if (salesRes.error) alert(salesRes.error.message);
    if (expenseRes.error) alert(expenseRes.error.message);
    if (depositRes.error) alert(depositRes.error.message);
    setRecords({ sales: salesRes.data || [], expenses: expenseRes.data || [], deposits: depositRes.data || [], credits: creditRes.data || [] });
    setLoading(false);
  }

  const cashSales = records.sales.filter((x) => x.payment_method === "CASH").reduce((a,x)=>a+number(x.sales_total),0);
  const creditPaidToday = records.credits.filter((x) => String(x.created_at || "").slice(0,10) === date).reduce((a,x)=>a+number(x.paid_amount),0);
  const cashExpenses = records.expenses.reduce((a,x)=>a+number(x.amount),0);
  const cashDeposits = records.deposits.filter((x) => String(x.method || "").toUpperCase() === "CASH").reduce((a,x)=>a+number(x.amount),0);
  const knownMovement = cashSales + creditPaidToday - cashExpenses - cashDeposits;
  const difference = actualCash === "" ? null : number(actualCash) - knownMovement;
  const debtPaymentsAll = records.credits.filter((x) => String(x.last_payment_at || "").slice(0,10) === date).reduce((a,x)=>a+number(x.paid_amount),0);

  function saveCount(value) {
    setActualCash(value);
    try { localStorage.setItem(`bless_cash_count_${businessId}_${date}`, value); } catch {}
  }

  function printClosing() {
    const win = window.open("", "_blank", "width=900,height=800");
    if (!win) { alert("Ruhusu pop-ups kisha ujaribu tena."); return; }
    win.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Daily Cash Closing</title><style>body{font-family:Arial;padding:28px;color:#111}h1{margin:0}.muted{color:#666}.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:20px}.card{border:1px solid #ddd;padding:12px}.card strong{display:block;font-size:18px;margin-top:5px}.total{font-size:22px;font-weight:800;margin-top:20px;padding:14px;border:2px solid #222}@media print{body{padding:0}}</style></head><body><h1>Bless Stationery</h1><div class="muted">Daily Cash Closing — ${escapeHtml(date)}</div><div class="grid"><div class="card">Cash Sales<strong>${escapeHtml(money(cashSales))}</strong></div><div class="card">Credit Paid Today<strong>${escapeHtml(money(creditPaidToday))}</strong></div><div class="card">Expenses<strong>${escapeHtml(money(cashExpenses))}</strong></div><div class="card">Cash Deposits<strong>${escapeHtml(money(cashDeposits))}</strong></div></div><div class="total">Known Cash Movement: ${escapeHtml(money(knownMovement))}<br>Actual Cash Count: ${escapeHtml(money(actualCash))}<br>Difference: ${escapeHtml(difference === null ? "-" : money(difference))}</div><p>Note: Mfumo unahesabu cash movement kutokana na records zilizopo. Debt payments za zamani zinahitaji payment history table ili ziweze kuhesabiwa kwa usahihi wa siku.</p><script>window.onload=()=>window.print()</script></body></html>`);
    win.document.close();
  }

  if (loading) return <PageLoading />;
  return (
    <div>
      <PageTitle title="Daily Cash Closing" subtitle="Funga siku, linganisha cash iliyotarajiwa na cash uliyo-count." />
      <div className="panel"><div className="form-grid"><Field label="Closing Date" type="date" value={date} onChange={setDate} /><Field label="Actual Cash Count" type="number" value={actualCash} onChange={saveCount} placeholder="Ingiza cash uliyo-count" /><div className="form-actions"><button className="secondary-btn" onClick={load}>↻ Refresh</button><button className="primary-btn" onClick={printClosing}>🖨 Print Closing</button></div></div></div>
      <div className="stats-grid">
        <StatCard title="Cash Sales" value={money(cashSales)} icon="💵" tone="blue" />
        <StatCard title="Credit Paid Today" value={money(creditPaidToday)} icon="💳" tone="green" />
        <StatCard title="Expenses" value={money(cashExpenses)} icon="💸" tone="red" />
        <StatCard title="Cash Deposits" value={money(cashDeposits)} icon="🏦" tone="purple" />
      </div>
      <div className="panel cash-closing-summary">
        <div><span>Known Cash Movement</span><strong>{money(knownMovement)}</strong></div>
        <div><span>Actual Cash Count</span><strong>{actualCash === "" ? "-" : money(actualCash)}</strong></div>
        <div><span>Difference</span><strong className={difference === null ? "" : difference === 0 ? "stock-ok" : "stock-danger"}>{difference === null ? "-" : money(difference)}</strong></div>
      </div>
      {debtPaymentsAll > 0 && <div className="notice">Kumbuka: kuna records zenye <strong>last_payment_at</strong> tarehe hii. Mfumo wa sasa hauna payment-history table, kwa hiyo hatutaki kuzidisha hesabu ya cash kwa kutumia paid_amount ya cumulative.</div>}
    </div>
  );
}

function OptionalToolsPage({ businessId }) {
  const [kind, setKind] = useState("suppliers");
  const [status, setStatus] = useState("");

  async function checkTable(tableName) {
    setStatus("Inakagua database...");
    const { error } = await supabase.from(tableName).select("id").eq("business_id", businessId).limit(1);
    if (error) setStatus(`Table '${tableName}' bado haijawekwa kwenye Supabase. UI iko tayari, lakini tunahitaji SQL ya database kabla ya kuitumia live.`);
    else setStatus(`Table '${tableName}' ipo na iko tayari kutumika.`);
  }

  const info = kind === "suppliers"
    ? { title: "Suppliers & Purchases", table: "suppliers", text: "Hifadhi suppliers, purchase orders, gharama za manunuzi na bidhaa zinazoingia." }
    : { title: "Audit Log", table: "audit_logs", text: "Hifadhi nani alifanya action gani, lini na kwenye record gani." };

  return (
    <div>
      <PageTitle title={info.title} subtitle="Module hii imeandaliwa bila kuvunja schema yako ya sasa." />
      <div className="panel">
        <div className="segmented-control"><button className={kind === "suppliers" ? "active" : ""} onClick={() => { setKind("suppliers"); setStatus(""); }}>Suppliers & Purchases</button><button className={kind === "audit" ? "active" : ""} onClick={() => { setKind("audit"); setStatus(""); }}>Audit Log</button></div>
        <div className="optional-module-card"><div className="optional-module-icon">{kind === "suppliers" ? "🚚" : "🛡"}</div><div><h3>{info.title}</h3><p>{info.text}</p><button className="primary-btn" onClick={() => checkTable(info.table)}>Check Database Setup</button></div></div>
        {status && <div className="notice" style={{ marginTop: 16 }}>{status}</div>}
        <div className="notice" style={{ marginTop: 16 }}><strong>Kwa nini sijaweka fake data?</strong><br />Module hizi zinahitaji tables mpya za Supabase. Sitaki App.jsx ianze ku-crash kwa sababu ya table ambazo hazipo. Ukiniambia tuendelee na module hii, nitakupa SQL yake kwanza, halafu tuiunganishe kabisa.</div>
      </div>
    </div>
  );
}

function PageTitle({ title, subtitle, button, onClick }) {
  return (
    <div className="page-title">
      <div>
        <h1>{title}</h1>
        <p>{subtitle}</p>
      </div>

      {button && (
        <button className="primary-btn" onClick={onClick}>
          {button}
        </button>
      )}
    </div>
  );
}

function Field({
  label,
  value,
  onChange,
  type = "text",
  placeholder = "",
  required = false,
}) {
  return (
    <div className="field">
      <label>{label}</label>
      <input
        type={type}
        value={value}
        placeholder={placeholder}
        required={required}
        onChange={(e) => onChange(e.target.value)}
      />
    </div>
  );
}

function SelectField({ label, value, onChange, options }) {
  return (
    <div className="field">
      <label>{label}</label>
      <select value={value} onChange={(e) => onChange(e.target.value)}>
        {options.map((x) => (
          <option value={x.value} key={x.value}>
            {x.label}
          </option>
        ))}
      </select>
    </div>
  );
}

function EmptyState({ text }) {
  return <div className="empty-state">{text}</div>;
}

function PageLoading() {
  return (
    <div className="page-loading">
      <div className="spinner" />
      <p>Inapakia data...</p>
    </div>
  );
}



function CustomerCreditsPage({ businessId, staff, refresh }) {
  const [rows, setRows] = useState([]);
  const [search, setSearch] = useState("");
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [useAmount, setUseAmount] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => { load(); }, [businessId, refresh]);

  async function load() {
    if (!businessId) return;
    setLoading(true);
    const { data, error } = await supabase
      .from("customer_credits")
      .select("*")
      .eq("business_id", businessId)
      .order("created_at", { ascending: false });
    if (error) alert(error.message);
    setRows(data || []);
    setLoading(false);
  }

  async function useCredit(row) {
    const amount = number(useAmount);
    if (amount <= 0 || amount > number(row.balance)) {
      alert("Kiasi cha Customer Credit si sahihi.");
      return;
    }
    setBusy(true);
    const used = number(row.used_amount) + amount;
    const balance = Math.max(0, number(row.original_amount) - used);
    const status = balance <= 0 ? "USED" : "PARTIAL";

    const { error: updateError } = await supabase
      .from("customer_credits")
      .update({ used_amount: used, status, updated_at: new Date().toISOString() })
      .eq("id", row.id)
      .eq("business_id", businessId);

    if (updateError) {
      alert(updateError.message);
      setBusy(false);
      return;
    }

    const { error: txnError } = await supabase.from("customer_credit_transactions").insert({
      business_id: businessId,
      customer_credit_id: row.id,
      transaction_type: "USE",
      amount,
      note: "Customer Credit imetumika",
      staff_id: staff?.id || null,
      created_by: staff?.id || null,
    });
    if (txnError) console.error(txnError);

    setSelected(null);
    setUseAmount("");
    setBusy(false);
    await load();
  }

  const available = rows.filter((x) => number(x.balance) > 0);
  const filtered = rows.filter((x) => {
    const q = search.trim().toLowerCase();
    return !q || String(x.customer_name || "").toLowerCase().includes(q) || String(x.customer_phone || "").includes(q);
  });
  const totalOriginal = rows.reduce((a, x) => a + number(x.original_amount), 0);
  const totalUsed = rows.reduce((a, x) => a + number(x.used_amount), 0);
  const totalBalance = rows.reduce((a, x) => a + number(x.balance), 0);

  if (loading) return <PageLoading />;

  return (
    <div>
      <PageTitle title="Customer Credits" subtitle="Change ya wateja na salio lao linaloweza kutumika kwenye manunuzi yajayo." />
      <div className="stats-grid small-stats">
        <StatCard title="Credit Zote" value={money(totalOriginal)} icon="💰" tone="blue" />
        <StatCard title="Zimetumika" value={money(totalUsed)} icon="↘" tone="purple" />
        <StatCard title="Credit Inayopatikana" value={money(totalBalance)} icon="✓" tone="green" />
        <StatCard title="Wateja Wenye Credit" value={new Set(available.map((x) => `${x.customer_name}|${x.customer_phone || ""}`)).size} icon="👥" tone="orange" />
      </div>

      <div className="panel">
        <div className="panel-header">
          <div><h3>Customer Credit Ledger</h3><span>Change inaingia hapa moja kwa moja baada ya cash sale yenye malipo yanayozidi total.</span></div>
          <button className="secondary-btn" onClick={load}>↻ Refresh</button>
        </div>
        <div className="toolbar">
          <input className="search-input" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Tafuta mteja au simu..." />
          <button className="secondary-btn small-btn" onClick={() => exportToExcel(filtered.map(x => ({
            Customer: x.customer_name, Phone: x.customer_phone || "", Original: x.original_amount, Used: x.used_amount, Balance: x.balance, Status: x.status, Date: formatDate(x.created_at)
          })), `Bless-Stationery-Customer-Credits-${todayDateInput()}.xlsx`, "Customer Credits")}>⬇ Export</button>
        </div>
        <div className="table-wrap">
          <table>
            <thead><tr><th>Mteja</th><th>Credit ya Mwanzo</th><th>Imetumika</th><th>Salio</th><th>Status</th><th>Tarehe</th><th>Action</th></tr></thead>
            <tbody>
              {filtered.map((x) => (
                <tr key={x.id}>
                  <td><strong>{x.customer_name}</strong><br/><small>{x.customer_phone || "-"}</small></td>
                  <td>{money(x.original_amount)}</td>
                  <td>{money(x.used_amount)}</td>
                  <td><strong className={number(x.balance) > 0 ? "stock-ok" : "muted"}>{money(x.balance)}</strong></td>
                  <td><span className={`status-chip ${String(x.status || "").toLowerCase()}`}>{x.status}</span></td>
                  <td>{formatDate(x.created_at)}</td>
                  <td>{number(x.balance) > 0 && <button className="secondary-btn small-btn" onClick={() => { setSelected(x); setUseAmount(""); }}>Tumia Credit</button>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {!filtered.length && <EmptyState text="Hakuna Customer Credit iliyopatikana." />}
      </div>

      {selected && (
        <div className="modal-backdrop" onClick={() => setSelected(null)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="panel-header">
              <div><h3>Tumia Customer Credit</h3><span>{selected.customer_name} — Salio {money(selected.balance)}</span></div>
              <button className="icon-btn" onClick={() => setSelected(null)}>×</button>
            </div>
            <Field label="Kiasi cha kutumia" type="number" value={useAmount} onChange={setUseAmount} min="0.01" max={number(selected.balance)} />
            <div className="notice" style={{ marginTop: 12 }}>Salio litakalobaki: <strong>{money(Math.max(0, number(selected.balance) - number(useAmount)))}</strong></div>
            <div className="form-actions" style={{ marginTop: 16 }}>
              <button className="secondary-btn" onClick={() => setSelected(null)}>Cancel</button>
              <button className="primary-btn" disabled={busy} onClick={() => useCredit(selected)}>{busy ? "Inahifadhi..." : "Thibitisha Kutumia"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function PayablesPage({ businessId, staff, refresh }) {
  const [rows, setRows] = useState([]);
  const [payments, setPayments] = useState([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [search, setSearch] = useState("");
  const [paying, setPaying] = useState(null);
  const [paymentAmount, setPaymentAmount] = useState("");
  const [form, setForm] = useState({
    creditor_name: "",
    creditor_phone: "",
    creditor_type: "SUPPLIER",
    original_amount: "",
    paid_amount: "0",
    due_date: "",
    note: "",
  });

  async function load() {
    setLoading(true);
    const [{ data, error }, { data: paymentRows }] = await Promise.all([
      supabase
        .from("payables")
        .select("*")
        .eq("business_id", businessId)
        .order("created_at", { ascending: false }),
      supabase
        .from("payable_payments")
        .select("*")
        .eq("business_id", businessId)
        .order("created_at", { ascending: false }),
    ]);

    if (error) alert(error.message);
    setRows(data || []);
    setPayments(paymentRows || []);
    setLoading(false);
  }

  useEffect(() => {
    if (businessId) load();
  }, [businessId, refresh]);

  async function addPayable(e) {
    e.preventDefault();

    const original = number(form.original_amount);
    const paid = number(form.paid_amount);

    if (!form.creditor_name.trim() || original <= 0) {
      alert("Weka jina la mtu/supplier na kiasi sahihi.");
      return;
    }

    if (paid < 0 || paid > original) {
      alert("Kiasi ulicholipa hakiwezi kuzidi deni.");
      return;
    }

    setBusy(true);

    const balance = Math.max(0, original - paid);
    const { data, error } = await supabase
      .from("payables")
      .insert({
        business_id: businessId,
        creditor_name: form.creditor_name.trim(),
        creditor_phone: form.creditor_phone.trim() || null,
        creditor_type: form.creditor_type,
        original_amount: original,
        paid_amount: paid,
        status: balance <= 0 ? "PAID" : paid > 0 ? "PARTIAL" : "UNPAID",
        due_date: form.due_date || null,
        note: form.note.trim() || null,
        staff_id: staff?.id || null,
        created_by: staff?.id || null,
      })
      .select()
      .single();

    if (error) {
      alert(error.message);
    } else {
      if (paid > 0 && data?.id) {
        await supabase.from("payable_payments").insert({
          business_id: businessId,
          payable_id: data.id,
          payment_date: todayDateInput(),
          amount: paid,
          method: "CASH",
          note: "Malipo ya mwanzo wakati wa kuingiza deni.",
          staff_id: staff?.id || null,
          created_by: staff?.id || null,
        });
      }

      setForm({
        creditor_name: "",
        creditor_phone: "",
        creditor_type: "SUPPLIER",
        original_amount: "",
        paid_amount: "0",
        due_date: "",
        note: "",
      });
      await load();
    }

    setBusy(false);
  }

  async function payPayable(row) {
    const pay = number(paymentAmount);

    if (pay <= 0 || pay > number(row.balance)) {
      alert("Kiasi cha malipo si sahihi.");
      return;
    }

    const paid = number(row.paid_amount) + pay;
    const balance = Math.max(0, number(row.original_amount) - paid);
    const status = balance <= 0 ? "PAID" : "PARTIAL";

    setBusy(true);

    const { error } = await supabase
      .from("payables")
      .update({
        paid_amount: paid,
        status,
        updated_at: new Date().toISOString(),
      })
      .eq("id", row.id)
      .eq("business_id", businessId);

    if (error) {
      alert(error.message);
      setBusy(false);
      return;
    }

    const { error: paymentError } = await supabase
      .from("payable_payments")
      .insert({
        business_id: businessId,
        payable_id: row.id,
        payment_date: todayDateInput(),
        amount: pay,
        method: "CASH",
        staff_id: staff?.id || null,
        created_by: staff?.id || null,
      });

    if (paymentError) alert(paymentError.message);

    setPaying(null);
    setPaymentAmount("");
    await load();
    setBusy(false);
  }

  const filtered = rows.filter((row) => {
    const q = search.trim().toLowerCase();
    if (!q) return true;
    return (
      String(row.creditor_name || "").toLowerCase().includes(q) ||
      String(row.creditor_phone || "").toLowerCase().includes(q) ||
      String(row.creditor_type || "").toLowerCase().includes(q)
    );
  });

  const totalOriginal = rows.reduce((a, x) => a + number(x.original_amount), 0);
  const totalPaid = rows.reduce((a, x) => a + number(x.paid_amount), 0);
  const totalBalance = rows.reduce((a, x) => a + number(x.balance), 0);

  return (
    <div>
      <PageTitle
        title="Madeni Yetu"
        subtitle="Watu, suppliers au huduma tunazodaiwa."
      />

      <div className="stats-grid small-stats">
        <StatCard title="Madeni Yote" value={money(totalOriginal)} icon="📌" tone="red" />
        <StatCard title="Tuliyolipa" value={money(totalPaid)} icon="💵" tone="green" />
        <StatCard title="Tunadaiwa" value={money(totalBalance)} icon="⚠️" tone="orange" />
        <StatCard title="Wanaotudai" value={rows.filter(x => number(x.balance) > 0).length} icon="👤" tone="purple" />
      </div>

      <div className="panel form-panel">
        <div className="panel-header">
          <div>
            <h3>Ongeza Mtu/Supplier Anayetudai</h3>
            <span>Weka deni tunalopaswa kulipa.</span>
          </div>
        </div>

        <form onSubmit={addPayable}>
          <div className="form-grid">
            <div className="field">
              <label>Jina *</label>
              <input
                value={form.creditor_name}
                onChange={(e) => setForm({ ...form, creditor_name: e.target.value })}
                placeholder="Mfano: ABC Supplies"
              />
            </div>

            <div className="field">
              <label>Simu</label>
              <input
                value={form.creditor_phone}
                onChange={(e) => setForm({ ...form, creditor_phone: e.target.value })}
                placeholder="07XXXXXXXX"
              />
            </div>

            <div className="field">
              <label>Aina</label>
              <select
                value={form.creditor_type}
                onChange={(e) => setForm({ ...form, creditor_type: e.target.value })}
              >
                <option value="SUPPLIER">Supplier</option>
                <option value="PERSON">Mtu</option>
                <option value="SERVICE">Huduma</option>
                <option value="OTHER">Nyingine</option>
              </select>
            </div>

            <div className="field">
              <label>Kiasi cha Deni *</label>
              <input
                type="number"
                min="0"
                value={form.original_amount}
                onChange={(e) => setForm({ ...form, original_amount: e.target.value })}
                placeholder="0"
              />
            </div>

            <div className="field">
              <label>Kiasi Tulicholipa</label>
              <input
                type="number"
                min="0"
                value={form.paid_amount}
                onChange={(e) => setForm({ ...form, paid_amount: e.target.value })}
              />
            </div>

            <div className="field">
              <label>Tarehe ya Kulipa</label>
              <input
                type="date"
                value={form.due_date}
                onChange={(e) => setForm({ ...form, due_date: e.target.value })}
              />
            </div>

            <div className="field full">
              <label>Maelezo</label>
              <textarea
                rows="2"
                value={form.note}
                onChange={(e) => setForm({ ...form, note: e.target.value })}
                placeholder="Maelezo ya deni..."
              />
            </div>

            <div className="field">
              <label>Salio Litakalobaki</label>
              <input
                value={money(Math.max(0, number(form.original_amount) - number(form.paid_amount)))}
                readOnly
              />
            </div>

            <div className="field full form-actions">
              <button className="primary-btn" disabled={busy}>
                {busy ? "Ina-save..." : "＋ Hifadhi Deni"}
              </button>
            </div>
          </div>
        </form>
      </div>

      <div className="panel">
        <div className="panel-header">
          <div>
            <h3>Orodha ya Wanaotudai</h3>
            <span>Madeni ambayo biashara yako inapaswa kulipa.</span>
          </div>
          <button
            className="secondary-btn small-btn"
            onClick={() =>
              exportToExcel(
                filtered.map((x) => ({
                  Jina: x.creditor_name,
                  Simu: x.creditor_phone || "",
                  Aina: x.creditor_type,
                  Deni: number(x.original_amount),
                  Tuliyolipa: number(x.paid_amount),
                  Salio: number(x.balance),
                  Status: x.status,
                  Tarehe: x.due_date || "",
                  Maelezo: x.note || "",
                })),
                "madeni-yetu.xlsx",
                "Madeni Yetu"
              )
            }
          >
            Export Excel
          </button>
        </div>

        <div className="toolbar">
          <input
            className="customer-search"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Tafuta jina, simu au aina..."
          />
        </div>

        {loading ? (
          <div className="empty-state">Inapakia...</div>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Mtu/Supplier</th>
                  <th>Aina</th>
                  <th>Deni</th>
                  <th>Tuliyolipa</th>
                  <th>Salio</th>
                  <th>Due Date</th>
                  <th>Status</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((row) => (
                  <tr key={row.id}>
                    <td>
                      <strong>{row.creditor_name}</strong>
                      {row.creditor_phone && <small>{row.creditor_phone}</small>}
                    </td>
                    <td>{row.creditor_type}</td>
                    <td>{money(row.original_amount)}</td>
                    <td>{money(row.paid_amount)}</td>
                    <td><strong>{money(row.balance)}</strong></td>
                    <td>{row.due_date || "-"}</td>
                    <td>
                      <span className={`status-chip ${row.status === "PAID" ? "success" : row.status === "PARTIAL" ? "warning" : "danger"}`}>
                        {row.status}
                      </span>
                    </td>
                    <td>
                      {number(row.balance) > 0 ? (
                        <button
                          className="primary-btn small-btn"
                          onClick={() => {
                            setPaying(row);
                            setPaymentAmount("");
                          }}
                        >
                          Lipa
                        </button>
                      ) : (
                        <span className="status-chip success">Imelipwa</span>
                      )}
                    </td>
                  </tr>
                ))}

                {!filtered.length && (
                  <tr>
                    <td colSpan="8">
                      <div className="empty-state">
                        <strong>Hakuna madeni yaliyopatikana</strong>
                        <p>Ongeza mtu au supplier anayetudai hapo juu.</p>
                      </div>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {paying && (
        <div className="modal-backdrop">
          <div className="modal-card" style={{ maxWidth: 430, margin: "8vh auto", background: "#fff", padding: 24 }}>
            <div className="panel-header">
              <div>
                <h3>Lipa Deni</h3>
                <span>{paying.creditor_name}</span>
              </div>
              <button className="icon-btn" onClick={() => setPaying(null)}>×</button>
            </div>

            <div className="mini-row">
              <span>Salio la sasa</span>
              <strong>{money(paying.balance)}</strong>
            </div>

            <div className="field" style={{ marginTop: 16 }}>
              <label>Kiasi cha kulipa</label>
              <input
                autoFocus
                type="number"
                min="0"
                max={number(paying.balance)}
                value={paymentAmount}
                onChange={(e) => setPaymentAmount(e.target.value)}
                placeholder="0"
              />
            </div>

            <div className="form-actions" style={{ marginTop: 18 }}>
              <button className="secondary-btn" onClick={() => setPaying(null)}>
                Cancel
              </button>
              <button className="primary-btn" disabled={busy} onClick={() => payPayable(paying)}>
                {busy ? "Inalipa..." : "Thibitisha Malipo"}
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="panel">
        <div className="panel-header">
          <div>
            <h3>Historia ya Malipo</h3>
            <span>Malipo yaliyofanyika kwenye madeni yetu.</span>
          </div>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Tarehe</th>
                <th>Deni</th>
                <th>Kiasi</th>
                <th>Method</th>
                <th>Reference</th>
                <th>Maelezo</th>
              </tr>
            </thead>
            <tbody>
              {payments.map((p) => {
                const debt = rows.find((r) => r.id === p.payable_id);
                return (
                  <tr key={p.id}>
                    <td>{p.payment_date || formatDate(p.created_at)}</td>
                    <td>{debt?.creditor_name || "-"}</td>
                    <td><strong>{money(p.amount)}</strong></td>
                    <td>{p.method || "CASH"}</td>
                    <td>{p.reference_no || "-"}</td>
                    <td>{p.note || "-"}</td>
                  </tr>
                );
              })}
              {!payments.length && (
                <tr>
                  <td colSpan="6">
                    <div className="empty-state">
                      <strong>Hakuna malipo bado</strong>
                      <p>Historia ya malipo itaonekana hapa.</p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

const CSS = `
* {
  box-sizing: border-box;
}

.bless-welcome-backdrop{position:fixed;inset:0;z-index:9999;display:grid;place-items:center;padding:24px;background:rgba(4,9,18,.78);backdrop-filter:blur(18px);animation:blessFade .35s ease}
.bless-welcome-card{position:relative;width:min(720px,94vw);overflow:hidden;border:1px solid rgba(255,255,255,.16);border-radius:30px;padding:46px;background:linear-gradient(145deg,rgba(20,31,56,.98),rgba(8,14,28,.98));box-shadow:0 35px 100px rgba(0,0,0,.55);color:#fff;text-align:center;animation:blessPop .5s cubic-bezier(.2,.9,.2,1)}
.bless-welcome-logo{width:74px;height:74px;margin:0 auto 14px;display:grid;place-items:center;border-radius:22px;font-size:36px;font-weight:900;background:linear-gradient(135deg,#5eead4,#60a5fa);color:#07111f;box-shadow:0 15px 45px rgba(96,165,250,.35)}
.bless-welcome-kicker{font-size:11px;letter-spacing:4px;font-weight:800;opacity:.68}.bless-welcome-card h1{margin:10px 0 8px;font-size:38px;line-height:1.05}.bless-welcome-sub{margin:0 auto 26px;color:rgba(255,255,255,.72);font-size:15px}.bless-pulse-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:20px 0 26px}.bless-pulse-grid div{padding:15px 10px;border:1px solid rgba(255,255,255,.09);background:rgba(255,255,255,.055);border-radius:17px}.bless-pulse-grid span{display:block;font-size:9px;letter-spacing:1.5px;opacity:.55}.bless-pulse-grid strong{display:block;margin-top:7px;font-size:14px}.bless-enter-btn{width:100%;padding:15px 18px;border:0;border-radius:16px;background:#fff;color:#07111f;font-weight:900;letter-spacing:.6px;cursor:pointer;transition:.2s}.bless-enter-btn:hover{transform:translateY(-2px);box-shadow:0 12px 30px rgba(255,255,255,.12)}.bless-enter-btn span{float:right;font-size:20px}.bless-welcome-card small{display:block;margin-top:16px;opacity:.45}.bless-welcome-card kbd,.bless-command kbd{padding:3px 7px;border-radius:6px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.12);font-family:inherit}.bless-welcome-orbit{position:absolute;width:220px;height:220px;border:1px solid rgba(96,165,250,.16);border-radius:50%;right:-120px;top:-120px}.bless-welcome-orbit.two{width:160px;height:160px;left:-90px;bottom:-80px;border-color:rgba(94,234,212,.13)}
.bless-command-backdrop{position:fixed;inset:0;z-index:9998;background:rgba(2,6,14,.62);backdrop-filter:blur(12px);display:flex;justify-content:center;align-items:flex-start;padding:10vh 20px}.bless-command{width:min(700px,96vw);border:1px solid rgba(255,255,255,.13);border-radius:24px;background:var(--panel-bg,#111827);box-shadow:0 30px 90px rgba(0,0,0,.5);overflow:hidden;animation:blessPop .25s ease}.bless-command-head{display:flex;justify-content:space-between;align-items:center;padding:22px 24px 12px}.bless-command-head span{font-size:10px;letter-spacing:2px;opacity:.5}.bless-command-head h3{margin:5px 0 0}.bless-command-head button{border:0;background:transparent;font-size:28px;opacity:.6;cursor:pointer}.bless-command-search{margin:0 20px 16px;padding:13px 15px;border:1px solid rgba(128,128,128,.2);border-radius:12px;display:flex;gap:10px;align-items:center;opacity:.72}.bless-command-search span{flex:1}.bless-command-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;padding:0 20px 22px}.bless-command-grid button{display:grid;grid-template-columns:28px 1fr 18px;align-items:center;gap:8px;text-align:left;padding:14px;border:1px solid rgba(128,128,128,.16);border-radius:14px;background:rgba(128,128,128,.06);cursor:pointer;color:inherit}.bless-command-grid button:hover{transform:translateY(-2px);border-color:rgba(96,165,250,.45);background:rgba(96,165,250,.09)}.bless-command-grid span{font-size:18px}.bless-command-grid strong{font-size:12px}.bless-command-grid em{opacity:.4}.bless-floating-command{position:fixed;right:22px;bottom:22px;z-index:1200;width:54px;height:54px;border-radius:18px;border:1px solid rgba(255,255,255,.16);background:linear-gradient(145deg,#17233d,#0c1426);color:#fff;box-shadow:0 12px 35px rgba(0,0,0,.28);cursor:pointer;display:grid;place-items:center}.bless-floating-command span{font-size:22px}.bless-floating-command small{font-size:7px;opacity:.55;position:absolute;bottom:5px}.bless-floating-command:hover{transform:translateY(-3px)}
@keyframes blessFade{from{opacity:0}to{opacity:1}}@keyframes blessPop{from{opacity:0;transform:translateY(18px) scale(.97)}to{opacity:1;transform:none}}


html,
body,
#root {
  margin: 0;
  min-height: 100%;
  font-family: Inter, Arial, sans-serif;
  background: #f4f7fb;
  color: #172033;
}

button,
input,
select {
  font: inherit;
}

button {
  cursor: pointer;
}

.loading-screen {
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  background: linear-gradient(135deg, #071d41, #0d5bd7);
}

.loading-box {
  text-align: center;
  color: white;
}

.logo-circle,
.brand-logo,
.mobile-logo {
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 22px;
  font-weight: 900;
  color: #0b4dcc;
  background: white;
}

.logo-circle {
  width: 75px;
  height: 75px;
  font-size: 38px;
  margin: auto;
}

.login-page {
  min-height: 100vh;
  display: grid;
  grid-template-columns: 1fr 1fr;
  background: #f5f8fc;
}

.login-brand {
  background: linear-gradient(145deg, #061b3d, #0a55c7);
  color: white;
  padding: 70px;
  display: flex;
  flex-direction: column;
  justify-content: center;
}

.brand-logo {
  width: 90px;
  height: 90px;
  font-size: 48px;
  margin-bottom: 25px;
}

.login-brand h1 {
  font-size: 45px;
  margin: 0 0 10px;
}

.login-brand p {
  font-size: 18px;
  opacity: .8;
}

.brand-features {
  margin-top: 45px;
  display: grid;
  gap: 18px;
  font-size: 17px;
}

.login-side {
  display: flex;
  overflow-y: auto;
  justify-content: center;
  align-items: center;
  padding: 30px;
}

.login-card {
  width: min(450px, 100%);
  background: white;
  padding: 42px;
  border-radius: 25px;
  box-shadow: 0 25px 70px rgba(10, 30, 70, .12);
}

.login-tabs {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 6px;
  background: #eef2f7;
  padding: 5px;
  border-radius: 12px;
  margin-bottom: 24px;
}

.login-tabs button {
  border: 0;
  background: transparent;
  color: #65738a;
  padding: 10px 12px;
  border-radius: 9px;
  font-weight: 800;
}

.login-tabs button.active {
  background: linear-gradient(135deg, #ffffff 0%, #eaf3ff 100%);
  color: #1262dc;
  box-shadow: 0 8px 22px rgba(0,0,0,.12);
  box-shadow: 0 2px 8px rgba(10, 30, 70, .08);
}

.success-box {
  background: #effaf2;
  color: #21743a;
  border: 1px solid #bde7c8;
  padding: 12px;
  border-radius: 10px;
  margin-bottom: 15px;
}

.register-switch {
  text-align: center;
  color: #718096;
  font-size: 13px;
  margin: 18px 0 0;
}

.register-switch button {
  border: 0;
  background: transparent;
  color: #1262dc;
  font-weight: 800;
  padding: 0;
}

.login-card h2 {
  font-size: 30px;
  margin: 0 0 8px;
}

.muted {
  color: #718096;
  margin-top: 0;
  margin-bottom: 28px;
}

.registration-section-title {
  margin-top: 22px;
  margin-bottom: 10px;
  font-size: 13px;
  font-weight: 900;
  color: #163c77;
  text-transform: uppercase;
  letter-spacing: .4px;
}

.plan-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 10px;
  margin-bottom: 6px;
}

.plan-card {
  border: 1px solid #d9e0eb;
  background: #fff;
  border-radius: 14px;
  padding: 13px 10px;
  text-align: left;
  color: #172033;
  min-height: 145px;
  display: flex;
  flex-direction: column;
  gap: 5px;
}

.plan-card:hover {
  border-color: #1970ff;
}

.plan-card.selected {
  border: 2px solid #1970ff;
  background: #f2f7ff;
  box-shadow: 0 0 0 3px rgba(25,112,255,.08);
}

.plan-price {
  font-size: 15px;
  font-weight: 900;
  color: #0b4dcc;
}

.plan-period {
  font-size: 12px;
  font-weight: 700;
  color: #52627a;
}

.plan-card small {
  line-height: 1.35;
  color: #718096;
}

.login-card select {
  width: 100%;
  border: 1px solid #d9e0eb;
  border-radius: 11px;
  padding: 13px 14px;
  outline: none;
  background: white;
  margin-bottom: 0;
}

.login-card select:focus {
  border-color: #1970ff;
  box-shadow: 0 0 0 3px rgba(25,112,255,.1);
}

.login-card select:disabled {
  background: #f3f5f8;
  color: #8b95a7;
}

.login-card label,
.field label {
  display: block;
  font-weight: 700;
  font-size: 13px;
  margin-bottom: 8px;
}

.login-card input,
.field input,
.field select {
  width: 100%;
  border: 1px solid #d9e0eb;
  border-radius: 11px;
  padding: 13px 14px;
  outline: none;
  background: white;
}

.login-card input:focus,
.field input:focus,
.field select:focus {
  border-color: #1970ff;
  box-shadow: 0 0 0 3px rgba(25,112,255,.1);
}

.login-card > label {
  margin-top: 18px;
}

.password-wrap {
  position: relative;
}

.password-wrap input {
  padding-right: 80px;
}

.show-password {
  position: absolute;
  right: 5px;
  top: 5px;
  bottom: 5px;
  border: 0;
  background: #edf3ff;
  color: #1260d7;
  border-radius: 8px;
  padding: 0 10px;
}

.panel-actions,
.report-toolbar {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
}

.secondary-btn {
  border: 1px solid #d7deea;
  background: #ffffff;
  color: #172033;
  border-radius: 10px;
  padding: 10px 14px;
  font-weight: 700;
  cursor: pointer;
  transition: .2s ease;
}

.secondary-btn:hover {
  background: #f3f6fb;
  border-color: #b9c5d8;
  transform: translateY(-1px);
}

.report-toolbar {
  justify-content: space-between;
  margin-bottom: 18px;
}

.primary-btn,
.secondary-btn {
  border: 0;
  border-radius: 11px;
  padding: 12px 18px;
  font-weight: 800;
}

.primary-btn {
  color: white;
  background: #1262dc;
  box-shadow: 0 6px 16px rgba(18,98,220,.18);
}

.primary-btn:hover {
  background: #0b50bd;
}

.secondary-btn {
  background: #eef2f7;
  color: #26364d;
}

.login-btn {
  width: 100%;
  margin-top: 28px;
  padding: 15px;
}

.login-footer {
  text-align: center;
  color: #8a94a6;
  font-size: 12px;
  margin: 25px 0 0;
}

.error-box {
  background: #fff0f0;
  color: #c42e2e;
  border: 1px solid #ffcaca;
  padding: 12px;
  border-radius: 10px;
  margin-bottom: 15px;
}

.mobile-logo {
  width: 60px;
  height: 60px;
  font-size: 30px;
  margin-bottom: 18px;
  display: none;
}

.app-shell {
  min-height: 100vh;
  display: flex;
}

.sidebar {
  width: 255px;
  flex-shrink: 0;
  background: linear-gradient(180deg, #061a35 0%, #0b2345 100%);
  color: white;
  min-height: 100vh;
  padding: 22px 14px;
  box-shadow: 8px 0 28px rgba(7,27,56,.10);
  display: flex;
  flex-direction: column;
}

.sidebar-brand {
  display: flex;
  align-items: center;
  gap: 11px;
  padding: 5px 8px 25px;
}

.small-logo {
  width: 43px;
  height: 43px;
  background: white;
  color: #1262dc;
  border-radius: 12px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 25px;
  font-weight: 900;
}

.sidebar-brand strong {
  display: block;
  font-size: 15px;
}

.sidebar-brand span {
  display: block;
  color: #9fb0ca;
  font-size: 11px;
  margin-top: 3px;
}

.sidebar-user {
  display: flex;
  gap: 10px;
  align-items: center;
  background: rgba(255,255,255,.06);
  padding: 11px;
  border-radius: 13px;
  margin-bottom: 18px;
}

.avatar,
.staff-avatar {
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 50%;
  font-weight: 900;
}

.avatar {
  width: 38px;
  height: 38px;
  background: #1262dc;
}

.sidebar-user strong {
  display: block;
  font-size: 13px;
}

.sidebar-user span {
  font-size: 10px;
  color: #9fb0ca;
}

.nav-title {
  color: #68809f;
  font-size: 10px;
  font-weight: 800;
  margin: 19px 10px 7px;
}

.nav-button {
  width: 100%;
  border: 0;
  background: transparent;
  color: #b6c3d6;
  padding: 11px 12px;
  border-radius: 10px;
  display: flex;
  gap: 12px;
  align-items: center;
  text-align: left;
  margin-bottom: 3px;
}

.nav-button:hover {
  background: rgba(255,255,255,.07);
  color: white;
}

.nav-button.active {
  background: #1262dc;
  color: white;
}

.nav-icon {
  width: 21px;
  text-align: center;
}

.logout-btn {
  margin-top: auto;
  border: 0;
  background: rgba(255,255,255,.06);
  color: #ffb5b5;
  padding: 12px;
  border-radius: 10px;
}

.main-area {
  flex: 1;
  min-width: 0;
}

.topbar {
  height: 72px;
  background: white;
  border-bottom: 1px solid #e4e9f1;
  padding: 0 28px;
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.topbar strong {
  display: block;
  font-size: 18px;
}

.topbar-date {
  color: #8490a3;
  font-size: 11px;
}

.top-actions {
  display: flex;
  gap: 12px;
  align-items: center;
}

.language-control {
  display: flex;
  align-items: center;
  gap: 6px;
  min-height: 38px;
  padding: 3px 9px;
  border: 1px solid #dce3ed;
  border-radius: 10px;
  background: #fff;
}

.language-icon { font-size: 16px; }
.language-label {
  font-size: 12px;
  font-weight: 700;
  color: #26364d;
}
.language-select {
  margin: 0 !important;
  border: 0;
  outline: 0;
  background: transparent;
  color: #26364d;
  font-size: 12px;
  min-width: 115px;
  cursor: pointer;
}
.google-translate-hidden {
  position: absolute !important;
  width: 1px !important;
  height: 1px !important;
  overflow: hidden !important;
  opacity: 0 !important;
  pointer-events: none !important;
}
.google-translate-hidden .goog-te-gadget,
.google-translate-hidden .goog-te-combo,
.google-translate-hidden .goog-logo-link,
.google-translate-hidden .goog-te-gadget span {
  display: none !important;
}

.refresh-btn {
  border: 1px solid #dce3ed;
  background: white;
  padding: 9px 12px;
  border-radius: 9px;
}

.user-pill {
  padding: 9px 13px;
  border-radius: 20px;
  background: #f3f6fa;
  font-size: 12px;
}

.online-dot {
  display: inline-block;
  width: 7px;
  height: 7px;
  background: #22b573;
  border-radius: 50%;
  margin-right: 5px;
}

.content {
  padding: 28px;
  max-width: 1700px;
  margin: auto;
}

.mobile-menu-btn {
  display: none;
  border: 0;
  background: #eef3f9;
  border-radius: 8px;
  padding: 8px 10px;
}

.welcome,
.page-title {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 20px;
  margin-bottom: 24px;
}

.welcome h1,
.page-title h1 {
  margin: 0;
  font-size: 27px;
}

.welcome p,
.page-title p {
  margin: 6px 0 0;
  color: #7c899c;
}

.quick-actions {
  display: flex;
  gap: 10px;
}

.stats-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 16px;
  margin-bottom: 18px;
}

.small-stats .stat-card {
  min-height: 105px;
}

.stat-card {
  min-height: 135px;
  background: white;
  border: 1px solid #e5eaf1;
  border-radius: 16px;
  padding: 19px;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.stat-title {
  display: block;
  color: #8490a3;
  font-size: 12px;
  margin-bottom: 9px;
}

.stat-card strong {
  display: block;
  font-size: 20px;
}

.stat-icon {
  width: 46px;
  height: 46px;
  border-radius: 13px;
  display: flex;
  justify-content: center;
  align-items: center;
  background: #edf4ff;
  font-size: 21px;
}

.stat-card.green .stat-icon {
  background: #e8f8ef;
}

.stat-card.purple .stat-icon {
  background: #f1eaff;
}

.stat-card.orange .stat-icon {
  background: #fff1df;
}

.stat-card.red .stat-icon {
  background: #ffeaea;
}

.dashboard-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 18px;
  margin-top: 18px;
}

.panel {
  background: white;
  border: 1px solid #e3e8f0;
  border-radius: 16px;
  padding: 20px;
  margin-bottom: 18px;
  box-shadow: 0 3px 14px rgba(30,50,80,.025);
}

.panel-header-actions {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
}

.small-btn {
  padding: 8px 12px;
  font-size: 12px;
}

.panel-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  margin-bottom: 17px;
}

.panel-header h3 {
  margin: 0 0 4px;
  font-size: 16px;
}

.panel-header span {
  color: #8a96a8;
  font-size: 11px;
}

.text-btn {
  color: #1262dc;
  border: 0;
  background: transparent;
  font-weight: 700;
}

.mini-row,
.report-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 13px 0;
  border-bottom: 1px solid #eef1f5;
}

.mini-row:last-child,
.report-row:last-child {
  border-bottom: 0;
}

.mini-row strong,
.report-row strong {
  display: block;
  font-size: 13px;
}

.mini-row span,
.report-row span {
  display: block;
  color: #929dae;
  font-size: 11px;
  margin-top: 3px;
}

.stock-warning {
  color: #d48700;
  font-weight: 800;
}

.stock-danger {
  color: #d93030;
  font-weight: 800;
}

.form-panel {
  margin-bottom: 18px;
}

.form-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
}

.field {
  min-width: 0;
}

.form-actions {
  display: flex;
  gap: 10px;
  align-items: center;
}

.full {
  grid-column: 1 / -1;
}

.toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 15px;
  margin-bottom: 16px;
}

.search-box {
  width: min(450px, 100%);
  border: 1px solid #dce3ed;
  border-radius: 10px;
  padding: 0 12px;
  display: flex;
  gap: 8px;
  align-items: center;
}

.search-box input {
  border: 0;
  outline: 0;
  width: 100%;
  padding: 11px 4px;
}

.record-count {
  color: #8390a4;
  font-size: 12px;
}

.table-wrap {
  width: 100%;
  overflow-x: auto;
}

table {
  width: 100%;
  border-collapse: collapse;
  min-width: 750px;
}

th {
  background: #f6f8fb;
  color: #68758a;
  text-align: left;
  font-size: 11px;
  padding: 12px;
  white-space: nowrap;
}

td {
  padding: 13px 12px;
  border-bottom: 1px solid #edf0f4;
  font-size: 12px;
  white-space: nowrap;
}

td strong {
  font-size: 13px;
}

.badge,
.payment-badge {
  display: inline-block;
  border-radius: 20px;
  padding: 5px 9px;
  font-size: 10px;
  font-weight: 800;
}

.badge.success {
  background: #e7f8ee;
  color: #1b9253;
}

.badge.warning {
  background: #fff5dd;
  color: #b57700;
}

.badge.danger {
  background: #ffe8e8;
  color: #d12e2e;
}

.payment-badge {
  background: #edf3ff;
  color: #2465c2;
}

.danger-small {
  border: 0;
  background: #fff0f0;
  color: #d12f2f;
  padding: 7px 10px;
  border-radius: 8px;
  font-size: 11px;
  font-weight: 700;
}

.empty-state {
  text-align: center;
  color: #8995a7;
  padding: 35px 15px;
  font-size: 13px;
}

.mode-switch,
.sale-type-tabs,
.report-tabs {
  display: flex;
  gap: 8px;
  margin-bottom: 16px;
  background: #eaf0f7;
  width: fit-content;
  padding: 5px;
  border-radius: 10px;
}

.mode-switch button,
.sale-type-tabs button,
.report-tabs button {
  border: 0;
  background: transparent;
  padding: 9px 13px;
  border-radius: 7px;
  color: #66758b;
  font-weight: 700;
  font-size: 12px;
}

.mode-switch button.selected,
.sale-type-tabs button.active,
.report-tabs button.active {
  background: white;
  color: #1262dc;
  box-shadow: 0 2px 6px rgba(0,0,0,.07);
}

.notice {
  padding: 12px 14px;
  background: #eef6ff;
  color: #2c619b;
  border-radius: 10px;
  font-size: 12px;
  margin-bottom: 15px;
}

.sale-layout {
  display: grid;
  grid-template-columns: 1fr 300px;
  gap: 18px;
}

.sale-preview {
  background: linear-gradient(145deg, #071d41, #1262dc);
  color: white;
  border: 0;
  min-height: 300px;
}

.preview-label {
  font-size: 10px;
  opacity: .7;
  letter-spacing: 1px;
}

.preview-total {
  font-size: 30px;
  font-weight: 900;
  margin: 30px 0;
}

.preview-line {
  display: flex;
  justify-content: space-between;
  padding: 11px 0;
  border-bottom: 1px solid rgba(255,255,255,.15);
}

.preview-line span {
  opacity: .7;
  font-size: 12px;
}

.preview-line strong {
  font-size: 12px;
}

.preview-hint {
  font-size: 11px;
  opacity: .65;
  margin-top: 25px;
  line-height: 1.6;
}

.sale-total {
  background: #f2f6fc;
  border-radius: 11px;
  padding: 16px;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.sale-total strong {
  color: #1262dc;
  font-size: 23px;
}

.sale-save {
  width: 100%;
  padding: 15px;
  font-size: 14px;
}

.profit-text {
  color: #19945a;
  font-weight: 700;
}

.service-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 16px;
}

.service-card {
  background: white;
  border: 1px solid #e2e8f0;
  border-radius: 16px;
  padding: 18px;
}

.service-icon {
  width: 42px;
  height: 42px;
  display: flex;
  align-items: center;
  justify-content: center;
  background: #edf4ff;
  color: #1262dc;
  border-radius: 12px;
}

.service-card h3 {
  font-size: 14px;
  min-height: 38px;
}

.service-card span {
  display: block;
  color: #8995a7;
  font-size: 11px;
}

.service-card > strong {
  display: block;
  font-size: 17px;
  margin: 5px 0 15px;
}

.summary-panel {
  text-align: center;
  padding: 30px;
}

.summary-panel h3 {
  margin: 0;
}

.net-number {
  font-size: 34px;
  font-weight: 900;
  color: #1262dc;
  margin: 10px;
}

.summary-panel p {
  color: #8490a3;
  font-size: 12px;
}

.staff-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 15px;
}

.staff-card {
  border: 1px solid #e2e7ef;
  padding: 16px;
  border-radius: 13px;
  display: flex;
  gap: 12px;
  align-items: center;
}

.staff-avatar {
  width: 45px;
  height: 45px;
  background: #eaf2ff;
  color: #1262dc;
}

.staff-card h3 {
  margin: 0 0 3px;
  font-size: 14px;
}

.staff-card span,
.staff-card small {
  display: block;
  font-size: 11px;
  color: #8995a7;
}

.staff-card small {
  color: #20a260;
  margin-top: 3px;
}

.page-loading {
  min-height: 300px;
  display: flex;
  align-items: center;
  justify-content: center;
  flex-direction: column;
  color: #8290a5;
}

.spinner {
  width: 30px;
  height: 30px;
  border: 3px solid rgba(255,255,255,.3);
  border-top-color: currentColor;
  border-radius: 50%;
  animation: spin 1s linear infinite;
}

.page-loading .spinner {
  color: #1262dc;
  border-color: #dce8fa;
  border-top-color: #1262dc;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

.modal-overlay {
  position: fixed;
  inset: 0;
  background: rgba(5,20,45,.65);
  display: flex;
  justify-content: center;
  align-items: center;
  padding: 20px;
  z-index: 1000;
  overflow: auto;
}

.receipt-modal {
  background: white;
  border-radius: 15px;
  padding: 15px;
  width: min(450px, 100%);
  box-shadow: 0 30px 80px rgba(0,0,0,.25);
}

.receipt-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  margin-bottom: 10px;
}

.receipt-print {
  width: 80mm;
  max-width: 100%;
  margin: auto;
  padding: 20px 12px;
  background: white;
  color: black;
  font-family: Arial, sans-serif;
}

.receipt-header {
  text-align: center;
}

.receipt-logo {
  width: 40px;
  height: 40px;
  border-radius: 9px;
  background: black;
  color: white;
  display: flex;
  justify-content: center;
  align-items: center;
  margin: auto;
  font-weight: 900;
  font-size: 22px;
}

.receipt-header h2 {
  margin: 8px 0 3px;
  font-size: 19px;
}

.receipt-header p {
  font-size: 10px;
  margin: 0;
}

.receipt-info {
  display: flex;
  justify-content: space-between;
  gap: 10px;
  margin: 15px 0;
}

.receipt-info span {
  display: block;
  font-size: 9px;
  color: #666;
}

.receipt-info strong {
  display: block;
  font-size: 10px;
  margin-top: 3px;
}

.receipt-divider {
  border-top: 1px dashed #777;
}

.receipt-item {
  padding: 15px 0;
  display: flex;
  justify-content: space-between;
  gap: 15px;
  font-size: 11px;
}

.receipt-total {
  display: flex;
  justify-content: space-between;
  padding: 15px 0;
  font-size: 15px;
}

.receipt-footer {
  text-align: center;
  margin-top: 25px;
}

.receipt-footer p {
  font-size: 10px;
}

.receipt-footer small {
  font-size: 9px;
  color: #777;
}

.print-note {
  font-size: 10px;
  text-align: center;
  color: #78869b;
  padding: 10px;
}

.mobile-overlay {
  display: none;
}


.eyebrow,
.panel-kicker {
  color: #1262dc;
  font-size: 10px;
  font-weight: 900;
  letter-spacing: 1.4px;
  margin-bottom: 5px;
}

.staff-page-title h1 {
  letter-spacing: -.5px;
}

.staff-stats {
  margin-bottom: 18px;
}

.staff-form-panel {
  border-color: #d8e5fb;
  background: linear-gradient(180deg, #ffffff, #fbfdff);
}

.staff-toolbar {
  align-items: center;
}

.toolbar-title {
  margin: 2px 0 0;
  font-size: 18px;
}

.staff-search {
  width: min(320px, 100%);
}

.staff-directory {
  display: grid;
  gap: 11px;
}

.staff-directory-card {
  display: grid;
  grid-template-columns: minmax(240px, 1fr) 150px auto;
  align-items: center;
  gap: 18px;
  padding: 15px;
  border: 1px solid #e5eaf2;
  border-radius: 14px;
  background: #fff;
  transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
}

.staff-directory-card:hover {
  transform: translateY(-1px);
  border-color: #cbdaf2;
  box-shadow: 0 8px 24px rgba(25, 57, 100, .07);
}

.staff-main {
  display: flex;
  align-items: center;
  gap: 13px;
}

.staff-avatar.admin-avatar {
  background: #e8efff;
  color: #174fae;
}

.staff-main h3 {
  margin: 0 0 7px;
  font-size: 14px;
}

.staff-meta {
  display: flex;
  align-items: center;
  gap: 7px;
  flex-wrap: wrap;
}

.role-chip,
.status-chip {
  display: inline-flex;
  align-items: center;
  border-radius: 999px;
  padding: 4px 8px;
  font-size: 9px;
  font-weight: 900;
  letter-spacing: .3px;
}

.role-chip.staff {
  background: #edf5ff;
  color: #2865af;
}

.role-chip.admin {
  background: #f0ebff;
  color: #6943bb;
}

.status-chip.active {
  background: #e8f8ef;
  color: #188c50;
}

.status-chip.inactive {
  background: #f1f3f6;
  color: #788497;
}

.staff-id-box {
  border-left: 1px solid #edf0f5;
  padding-left: 18px;
}

.staff-id-box span {
  display: block;
  color: #8995a7;
  font-size: 9px;
  margin-bottom: 4px;
}

.staff-id-box strong {
  font-size: 11px;
  letter-spacing: .7px;
}

.staff-actions {
  display: flex;
  justify-content: flex-end;
  gap: 7px;
}

.status-btn {
  border: 0;
  border-radius: 8px;
  padding: 8px 10px;
  font-size: 10px;
  font-weight: 800;
}

.status-btn.deactivate {
  background: #fff5e8;
  color: #a86500;
}

.status-btn.activate {
  background: #e8f8ef;
  color: #188c50;
}



.saas-toolbar {
  display: grid;
  grid-template-columns: minmax(240px, 1fr) 180px 190px auto;
  gap: 10px;
  margin-bottom: 18px;
}

.saas-toolbar input,
.saas-toolbar select {
  width: 100%;
  border: 1px solid #dbe3ef;
  border-radius: 10px;
  padding: 11px 12px;
  background: #fff;
  font: inherit;
}

.saas-table-wrap {
  width: 100%;
  overflow-x: auto;
}

.saas-table {
  width: 100%;
  border-collapse: collapse;
  min-width: 850px;
}

.saas-table th,
.saas-table td {
  padding: 13px 12px;
  text-align: left;
  border-bottom: 1px solid #edf1f7;
  white-space: nowrap;
}

.saas-table th {
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: .04em;
  color: #718096;
  background: #f8fafc;
}

.saas-table tbody tr {
  cursor: pointer;
  transition: background .15s ease;
}

.saas-table tbody tr:hover {
  background: #f8fbff;
}

.saas-status {
  display: inline-flex;
  align-items: center;
  border-radius: 999px;
  padding: 5px 9px;
  font-size: 11px;
  font-weight: 700;
  letter-spacing: .02em;
}

.saas-status.active {
  background: #dcfce7;
  color: #166534;
}

.saas-status.trialing {
  background: #ede9fe;
  color: #6d28d9;
}

.saas-status.pending-payment {
  background: #fff7ed;
  color: #c2410c;
}

.saas-status.expired,
.saas-status.cancelled {
  background: #fee2e2;
  color: #b91c1c;
}

.saas-detail-panel {
  margin-top: 18px;
}

.saas-detail-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 12px;
}

.saas-detail-grid > div {
  padding: 13px;
  border: 1px solid #edf1f7;
  border-radius: 10px;
  background: #fbfcfe;
}

.saas-detail-grid span {
  display: block;
  color: #718096;
  font-size: 12px;
  margin-bottom: 5px;
}

.saas-detail-grid strong {
  display: block;
  word-break: break-word;
}

.saas-subtitle {
  margin: 22px 0 10px;
}

.platform-nav-title {
  margin-top: 12px;
  color: #7c3aed;
}

.debt-payment-box { min-width: 230px; padding: 10px; border: 1px solid var(--border, #d9dee7); border-radius: 12px; background: var(--surface-2, #f7f9fc); }
.debt-payment-label { font-size: 12px; font-weight: 700; margin-bottom: 6px; }
.debt-payment-box input { width: 100%; box-sizing: border-box; }
.debt-payment-balance { margin: 7px 0; font-size: 12px; color: var(--muted, #667085); }
.debt-payment-actions { display:flex; gap:6px; flex-wrap:wrap; }
@media (max-width: 900px) {
  .saas-toolbar {
    grid-template-columns: 1fr 1fr;
  }

  .saas-detail-grid {
    grid-template-columns: 1fr 1fr;
  }
  .staff-directory-card {
    grid-template-columns: 1fr;
  }

  .staff-id-box {
    border-left: 0;
    border-top: 1px solid #edf0f5;
    padding: 10px 0 0;
  }

  .staff-actions {
    justify-content: flex-start;
  }
}

@media print {
  body * {
    visibility: hidden !important;
  }

  .receipt-print,
  .receipt-print * {
    visibility: visible !important;
  }

  .receipt-print {
    position: absolute;
    left: 0;
    top: 0;
    width: 80mm;
    margin: 0;
    box-shadow: none;
  }

  .no-print {
    display: none !important;
  }

  @page {
    size: 80mm auto;
    margin: 0;
  }
}

@media (max-width: 1100px) {
  .stats-grid {
    grid-template-columns: repeat(2, 1fr);
  }

  .service-grid {
    grid-template-columns: repeat(2, 1fr);
  }

  .form-grid {
    grid-template-columns: repeat(2, 1fr);
  }
}

@media (max-width: 850px) {
  .login-page {
    grid-template-columns: 1fr;
  }

  .login-brand {
    display: none;
  }

  .mobile-logo {
    display: flex;
  }

  .sidebar {
    position: fixed;
    left: -270px;
    top: 0;
    bottom: 0;
    z-index: 100;
    transition: left .2s;
  }

  .sidebar.open {
    left: 0;
  }

  .mobile-overlay {
    display: block;
    position: fixed;
    inset: 0;
    background: rgba(0,0,0,.4);
    z-index: 90;
  }

  .mobile-menu-btn {
    display: block;
    margin-right: 12px;
  }

  .topbar {
    padding: 0 15px;
  }

  .topbar {
    justify-content: flex-start;
    gap: 8px;
  }

  .top-actions {
    margin-left: auto;
  }

  .user-pill {
    display: none;
  }

  .content {
    padding: 18px;
  }

  .dashboard-grid,
  .sale-layout {
    grid-template-columns: 1fr;
  }

  .staff-grid {
    grid-template-columns: 1fr 1fr;
  }
}

@media (max-width: 600px) {
  .saas-toolbar {
    grid-template-columns: 1fr;
  }

  .saas-detail-grid {
    grid-template-columns: 1fr;
  }
  .stats-grid,
  .form-grid,
  .service-grid,
  .staff-grid,
  .plan-grid {
    grid-template-columns: 1fr;
  }

  .welcome,
  .page-title {
    align-items: flex-start;
    flex-direction: column;
  }

  .quick-actions {
    width: 100%;
  }

  .quick-actions button {
    flex: 1;
  }

  .topbar-date {
    display: none;
  }

  .refresh-btn {
    padding: 8px;
    font-size: 11px;
  }

  .login-card {
    padding: 28px 22px;
  }

  .mode-switch,
  .sale-type-tabs,
  .report-tabs {
    width: 100%;
    overflow-x: auto;
  }

  .mode-switch button,
  .sale-type-tabs button,
  .report-tabs button {
    white-space: nowrap;
  }
}

.subscription-gate {
  max-width: 620px;
  margin: 60px auto;
  text-align: center;
  padding: 42px 28px;
}

.subscription-gate-icon {
  font-size: 46px;
  margin-bottom: 12px;
}

.subscription-gate h2 {
  margin: 0 0 10px;
}

.subscription-gate p {
  line-height: 1.6;
}

.subscription-gate-actions {
  margin-top: 24px;
}

/* Dashboard polish layer */
.content { max-width: 1600px; margin: 0 auto; width: 100%; }
.welcome { position: relative; overflow: hidden; border: 1px solid #e7eef8; box-shadow: 0 14px 40px rgba(15, 42, 78, .07); }
.welcome::after { content: ""; position: absolute; width: 220px; height: 220px; border-radius: 50%; right: -70px; top: -100px; background: rgba(18,98,220,.08); pointer-events: none; }
.stats-grid { gap: 16px; }
.stat-card { border: 1px solid #e8eef6; box-shadow: 0 10px 28px rgba(15,42,78,.055); transition: transform .18s ease, box-shadow .18s ease; }
.stat-card:hover { transform: translateY(-3px); box-shadow: 0 16px 34px rgba(15,42,78,.10); }
.panel, .table-wrap, .form-card, .chart-card { border: 1px solid #e7edf6; box-shadow: 0 10px 30px rgba(15,42,78,.055); }
.primary-btn, .secondary-btn, .refresh-btn, button { transition: transform .15s ease, box-shadow .15s ease, opacity .15s ease; }
.primary-btn:hover, .secondary-btn:hover, .refresh-btn:hover { transform: translateY(-1px); }
.topbar { backdrop-filter: blur(12px); box-shadow: 0 1px 0 rgba(15,42,78,.06); }
@media (max-width: 900px) {
  .content { padding-bottom: 28px; }
  .welcome { border-radius: 18px; }
}


/* Final quality layer */
.toolbar-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.row-actions{display:flex;gap:6px;align-items:center;flex-wrap:wrap}
.settings-grid{display:grid;grid-template-columns:minmax(0,1.4fr) minmax(280px,.8fr);gap:18px}
.settings-preview{text-align:center;display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:320px}
.settings-preview h2{margin:14px 0 4px}
.settings-preview p{margin:3px 0;color:#667085}
.settings-logo img{width:72px;height:72px;object-fit:contain;border-radius:14px;border:1px solid #e5e7eb}
.status-chip.overdue{background:#fff1f0;color:#b42318}
.status-chip.paid{background:#ecfdf3;color:#027a48}
.status-chip.partial{background:#fffaeb;color:#b54708}
.status-chip.unpaid{background:#f2f4f7;color:#344054}
.search-box input{min-width:150px}
input:focus,select:focus{outline:none;border-color:#1262dc;box-shadow:0 0 0 3px rgba(18,98,220,.10)}
button:disabled{cursor:not-allowed;opacity:.55;transform:none!important}
.table-wrap{overflow-x:auto}
.table-wrap table{min-width:760px}
.mini-row{transition:background .15s ease}
.mini-row:hover{background:#f8fbff}
.notice{border-radius:12px}
@media (max-width:900px){
  .settings-grid{grid-template-columns:1fr}
  .toolbar-actions{width:100%}
  .toolbar-actions select,.toolbar-actions button{flex:0 0 auto}
}
@media print{
  .sidebar,.topbar,.no-print,.mobile-overlay{display:none!important}
  .main-area{margin:0!important}
  .content{max-width:none!important;padding:0!important}
}

/* ===== BLESS STATIONERY DISPLAY / UI MASTER POLISH ===== */
:root{
  --bs-primary:#1262dc;
  --bs-primary-dark:#0b4fb8;
  --bs-navy:#071d3d;
  --bs-bg:#f6f8fc;
  --bs-card:#ffffff;
  --bs-border:#e6ebf2;
  --bs-text:#172033;
  --bs-muted:#738198;
  --bs-success:#079455;
  --bs-warning:#d98b00;
  --bs-danger:#d92d20;
  --bs-radius:16px;
  --bs-shadow:0 10px 30px rgba(16,38,72,.06);
}

body{
  background:
    radial-gradient(circle at 90% 0%,rgba(18,98,220,.055),transparent 28%),
    var(--bs-bg);
  color:var(--bs-text);
  -webkit-font-smoothing:antialiased;
}

*{scrollbar-width:thin;scrollbar-color:#c7d1df transparent}
*::-webkit-scrollbar{width:8px;height:8px}
*::-webkit-scrollbar-thumb{background:#c7d1df;border-radius:99px}
*::-webkit-scrollbar-track{background:transparent}

.app-shell{background:var(--bs-bg)}
.sidebar{
  position:sticky;
  top:0;
  height:100vh;
  overflow-y:auto;
  border-right:1px solid rgba(255,255,255,.04);
  box-shadow:12px 0 35px rgba(7,29,61,.08);
}
.sidebar-brand{position:sticky;top:0;z-index:2;background:linear-gradient(180deg,#061a35 0%,#061a35 82%,rgba(6,26,53,0) 100%)}
.nav-button{font-size:13px;transition:background .18s ease,color .18s ease,transform .18s ease}
.nav-button:hover{transform:translateX(2px)}
.nav-button.active{box-shadow:0 8px 18px rgba(18,98,220,.22)}

.main-area{background:var(--bs-bg)}
.topbar{
  position:sticky;
  top:0;
  z-index:20;
  background:rgba(255,255,255,.92);
}
.topbar strong{letter-spacing:-.2px}
.content{padding:30px;max-width:1680px}

.welcome{
  background:linear-gradient(135deg,#fff 0%,#f5f9ff 100%);
  border-radius:20px;
  padding:24px;
  min-height:118px;
}
.welcome h1,.page-title h1{letter-spacing:-.6px}
.page-title{padding-bottom:4px}

.stats-grid{grid-template-columns:repeat(4,minmax(0,1fr))}
.stat-card{
  position:relative;
  overflow:hidden;
  min-width:0;
  border-radius:18px;
  background:rgba(255,255,255,.96);
}
.stat-card::after{
  content:"";
  position:absolute;
  right:-35px;
  bottom:-45px;
  width:120px;
  height:120px;
  border-radius:50%;
  background:rgba(18,98,220,.035);
}
.stat-card strong{font-size:22px;letter-spacing:-.35px}
.stat-title{text-transform:none}
.stat-icon{position:relative;z-index:1;box-shadow:0 5px 14px rgba(18,98,220,.08)}

.panel,.table-wrap,.form-card,.chart-card{
  border-radius:18px;
  box-shadow:var(--bs-shadow);
  background:var(--bs-card);
}
.panel{padding:21px}
.panel-header{align-items:flex-start}
.panel-header h3{letter-spacing:-.15px}

.form-panel{border-radius:18px}
.form-grid{align-items:end}
.field label{color:#344054}
.field input,.field select,.login-card input,.login-card select{
  min-height:45px;
  transition:border-color .16s ease,box-shadow .16s ease,background .16s ease;
}
.field input:hover,.field select:hover{border-color:#bcc8d8}
input::placeholder{color:#a0aabc}

.primary-btn,.secondary-btn,.refresh-btn,.small-btn,.login-btn{
  transition:transform .16s ease,box-shadow .16s ease,background .16s ease,border-color .16s ease;
}
.primary-btn:hover{box-shadow:0 9px 22px rgba(18,98,220,.22)}
.primary-btn:active,.secondary-btn:active,.refresh-btn:active{transform:translateY(0) scale(.98)}
.secondary-btn{border:1px solid #d9e1ec}
.secondary-btn:hover{background:#f7f9fc;border-color:#c5cfdd}
.text-btn{border-radius:8px;padding:7px 9px}
.text-btn:hover{background:#edf4ff}

.toolbar{
  background:#fff;
  border:1px solid var(--bs-border);
  border-radius:16px;
  padding:12px;
  box-shadow:0 6px 20px rgba(16,38,72,.035);
}
.toolbar .search-box{background:#fff}
.toolbar-actions{gap:8px}

.table-wrap{
  border:1px solid var(--bs-border);
  overflow:auto;
}
.table-wrap table{
  width:100%;
  border-collapse:separate;
  border-spacing:0;
}
.table-wrap thead th{
  position:sticky;
  top:0;
  z-index:1;
  background:#f8fafc;
  color:#667085;
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.45px;
  white-space:nowrap;
}
.table-wrap tbody td{
  background:#fff;
  border-bottom:1px solid #eef1f5;
  vertical-align:middle;
}
.table-wrap tbody tr:last-child td{border-bottom:0}
.table-wrap tbody tr:hover td{background:#fbfdff}
.row-actions button{white-space:nowrap}

.status-chip{
  display:inline-flex;
  align-items:center;
  justify-content:center;
  gap:5px;
  min-height:25px;
  padding:4px 9px;
  border-radius:999px;
  font-size:11px;
  font-weight:800;
  border:1px solid transparent;
}

.notice,.error-box,.success-box{
  box-shadow:0 5px 16px rgba(16,38,72,.035);
}

.modal-backdrop,.modal-overlay{
  backdrop-filter:blur(4px);
}
.modal,.modal-card,.dialog{
  border-radius:20px!important;
  box-shadow:0 24px 70px rgba(7,29,61,.20)!important;
  border:1px solid #e5eaf1!important;
}

.empty-state{
  padding:42px 20px;
  text-align:center;
  color:#7b8799;
}
.empty-state strong{display:block;color:#344054;margin-bottom:5px}
.empty-state p{margin:0;font-size:13px}

.loading-screen{
  background:
    radial-gradient(circle at 20% 20%,rgba(255,255,255,.08),transparent 30%),
    linear-gradient(145deg,#061a35,#1262dc);
}
.loading-box{background:rgba(255,255,255,.07);padding:36px;border-radius:24px;backdrop-filter:blur(10px)}

.login-page{background:#f6f8fc}
.login-brand{
  position:relative;
  overflow:hidden;
  background:
    radial-gradient(circle at 85% 15%,rgba(66,145,255,.22),transparent 30%),
    linear-gradient(145deg,#061a35,#0b57c7);
}
.login-brand::after{
  content:"";
  position:absolute;
  width:360px;height:360px;
  border-radius:50%;
  right:-150px;bottom:-160px;
  border:55px solid rgba(255,255,255,.045);
}
.login-card{border:1px solid #e6ebf2}
.login-card h2{letter-spacing:-.7px}

.quick-actions{flex-wrap:wrap}
.panel-header-actions{justify-content:flex-end}

.mobile-menu-btn{
  box-shadow:0 4px 12px rgba(16,38,72,.08);
}
.mobile-overlay{
  backdrop-filter:blur(3px);
}

@media (max-width:1200px){
  .stats-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
  .dashboard-grid{grid-template-columns:1fr}
  .form-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
  .content{padding:24px}
}
@media (max-width:900px){
  .app-shell{display:block}
  .sidebar{
    position:fixed;
    left:0;
    top:0;
    z-index:100;
    width:280px;
    transform:translateX(-105%);
    transition:transform .22s ease;
    box-shadow:18px 0 45px rgba(7,29,61,.25);
  }
  .sidebar.open{transform:translateX(0)}
  .main-area{width:100%}
  .topbar{
    height:64px;
    padding:0 14px;
  }
  .mobile-menu-btn{display:inline-flex}
  .topbar-date{display:none}
  .topbar strong{font-size:15px}
  .content{padding:18px 14px 28px}
  .welcome,.page-title{align-items:flex-start;flex-direction:column}
  .welcome{padding:19px}
  .welcome h1,.page-title h1{font-size:23px}
  .quick-actions{width:100%}
  .quick-actions button{flex:1 1 auto}
  .stats-grid{grid-template-columns:1fr 1fr;gap:11px}
  .stat-card{min-height:115px;padding:15px}
  .stat-card strong{font-size:18px}
  .stat-icon{width:40px;height:40px}
  .toolbar{align-items:stretch;flex-direction:column}
  .toolbar .search-box{width:100%}
  .toolbar-actions{width:100%}
  .toolbar-actions>*{flex:1 1 auto}
  .form-grid{grid-template-columns:1fr}
  .full{grid-column:auto}
  .settings-grid{grid-template-columns:1fr}
  .panel{padding:16px}
}
@media (max-width:620px){
  .login-page{display:block}
  .login-brand{display:none}
  .login-side{min-height:100vh;padding:16px}
  .login-card{padding:25px 19px;border-radius:20px}
  .mobile-logo{display:flex;margin:0 auto 15px}
  .plan-grid{grid-template-columns:1fr}
  .stats-grid{grid-template-columns:1fr}
  .stat-card{min-height:100px}
  .top-actions .refresh-btn{display:none}
  .user-pill{max-width:135px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .panel-header{flex-direction:column}
  .panel-header-actions{width:100%;justify-content:flex-start}
  .panel-header-actions button{flex:1 1 auto}
  .report-toolbar{align-items:stretch;flex-direction:column}
  .report-toolbar>*{width:100%}
  .search-box{width:100%}
  .content{padding-left:10px;padding-right:10px}
  .table-wrap{margin-left:-2px;margin-right:-2px;border-radius:14px}
}
@media print{
  body{background:#fff!important}
  .sidebar,.topbar,.mobile-overlay,.mobile-menu-btn,.no-print{display:none!important}
  .main-area{margin:0!important;width:100%!important}
  .content{max-width:none!important;padding:0!important}
  .panel,.table-wrap{box-shadow:none!important;border:1px solid #ddd!important}
}


/* Settings, themes and account security */
.app-shell{
  --brand:#0f3d5e;
  --brand-2:#0b5f7a;
  --accent:#0ea5a4;
  --accent-soft:#e6fffb;
  --page-bg:#f4f7fb;
  --panel-bg:#ffffff;
  --text:#172033;
  --muted:#718096;
  --border:#e5eaf1;
  --sidebar-text:#dbeafe;
}
.app-shell.theme-emerald{--brand:#064e3b;--brand-2:#047857;--accent:#10b981;--accent-soft:#e7fff5}
.app-shell.theme-purple{--brand:#3b1f6f;--brand-2:#5b21b6;--accent:#8b5cf6;--accent-soft:#f1eaff}
.app-shell.theme-slate{--brand:#172033;--brand-2:#334155;--accent:#06b6d4;--accent-soft:#e6fbff}
.app-shell.theme-rose{--brand:#5f172b;--brand-2:#9f1239;--accent:#e11d48;--accent-soft:#fff0f3}
.app-shell .sidebar{background:linear-gradient(180deg,var(--brand),var(--brand-2))}
.app-shell .nav-button.active{background:var(--accent);color:#fff;box-shadow:0 7px 18px color-mix(in srgb,var(--accent) 30%,transparent)}
.app-shell .primary-btn{background:var(--accent);border-color:var(--accent)}
.app-shell .primary-btn:hover{filter:brightness(.95)}
.app-shell .stat-icon{color:var(--accent)}
.app-shell .content{background:var(--page-bg)}
.app-shell .topbar{border-bottom-color:var(--border)}
.app-shell.density-compact .panel{padding:16px}
.app-shell.density-compact .table-wrap th,.app-shell.density-compact .table-wrap td{padding:8px 10px}
.app-shell.density-compact .stat-card{min-height:108px;padding:14px}

.settings-section-title{margin:28px 0 12px}
.settings-section-title h2{margin:0 0 5px;font-size:20px;color:var(--text)}
.settings-section-title p{margin:0;color:var(--muted);font-size:13px}
.theme-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px;margin-bottom:22px}
.theme-card{position:relative;text-align:left;border:1px solid var(--border);background:#fff;border-radius:16px;padding:12px;transition:.18s;box-shadow:0 5px 18px rgba(16,38,72,.04)}
.theme-card:hover{transform:translateY(-2px);box-shadow:0 10px 25px rgba(16,38,72,.08)}
.theme-card.selected{border:2px solid var(--accent);padding:11px;box-shadow:0 10px 25px rgba(16,38,72,.10)}
.theme-preview{height:62px;border-radius:11px;background:var(--preview-main);padding:10px;display:flex;gap:6px;align-items:flex-end;overflow:hidden;margin-bottom:11px}
.theme-preview span{display:block;border-radius:5px;background:#fff;height:18px;flex:1;opacity:.92}
.theme-preview span:first-child{height:38px;background:var(--preview-accent)}
.theme-preview span:last-child{height:27px;opacity:.45}
.theme-card-copy strong{display:block;color:#1f2937;font-size:13px}
.theme-card-copy small{display:block;color:#7b8799;margin-top:3px;font-size:11px}
.theme-check{position:absolute;right:10px;top:10px;width:24px;height:24px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:var(--accent);color:#fff;font-weight:900}
.preference-row{display:flex;justify-content:space-between;gap:20px;align-items:center;padding:5px 0 16px}
.preference-row p{margin:4px 0 0;color:var(--muted);font-size:12px}
.segmented-control{display:flex;background:#f0f3f7;border-radius:10px;padding:3px;flex-shrink:0}
.segmented-control button{border:0;background:transparent;border-radius:8px;padding:8px 12px;color:#667085;font-size:12px;font-weight:700}
.segmented-control button.active{background:#fff;color:var(--accent);box-shadow:0 2px 7px rgba(16,38,72,.10)}
.settings-grid-bottom{margin-top:0}
.password-settings-form{display:grid;gap:11px}
.password-settings-field{margin:0}
.password-hint{font-size:11px;color:#7b8799;line-height:1.5;margin:0 0 3px}
.settings-help-panel{margin-top:18px}
.settings-help-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
.settings-help-grid>div{border:1px solid var(--border);background:#fafbfd;border-radius:13px;padding:14px}
.settings-help-grid strong{display:block;font-size:13px;color:#344054;margin-bottom:5px}
.settings-help-grid span{display:block;font-size:11px;line-height:1.5;color:#7b8799}

@media (max-width:1100px){.theme-grid{grid-template-columns:repeat(3,minmax(0,1fr))}.settings-help-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:700px){.theme-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.preference-row{align-items:flex-start;flex-direction:column}.segmented-control{width:100%}.segmented-control button{flex:1}.settings-help-grid{grid-template-columns:1fr}}
@media (max-width:480px){.theme-grid{grid-template-columns:1fr}}


.customer-search{width:100%;max-width:520px;margin:0 0 16px;padding:12px 14px;border:1px solid var(--border);border-radius:10px;background:#fff;color:#1f2937}
.customer-detail-panel{margin-top:18px}
.analytics-filter .form-grid{align-items:end}
.analytics-bars{display:grid;gap:14px;padding-top:8px}
.analytics-bar-row{display:grid;grid-template-columns:52px 1fr 120px;gap:10px;align-items:center;font-size:12px}
.analytics-bar-track{height:12px;background:#edf2f7;border-radius:999px;overflow:hidden}
.analytics-bar-fill{height:100%;border-radius:999px;background:var(--accent);min-width:2px}
.cash-closing-summary{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px;margin-top:18px}
.cash-closing-summary>div{padding:18px;border:1px solid var(--border);border-radius:14px;background:#fff}
.cash-closing-summary span{display:block;color:var(--muted);font-size:12px;margin-bottom:7px}
.cash-closing-summary strong{font-size:22px;color:#1f2937}
.optional-module-card{display:flex;gap:18px;align-items:flex-start;margin-top:20px;padding:20px;border:1px solid var(--border);border-radius:16px;background:#fafbfd}
.optional-module-card h3{margin:0 0 7px}
.optional-module-card p{margin:0 0 14px;color:var(--muted);line-height:1.6}
.optional-module-icon{width:54px;height:54px;border-radius:14px;display:grid;place-items:center;background:#eef4ff;font-size:25px;flex:0 0 auto}
.stock-ok{color:#067647}
@media (max-width:800px){.analytics-bar-row{grid-template-columns:42px 1fr 90px}.cash-closing-summary{grid-template-columns:1fr}.optional-module-card{flex-direction:column}}

.payables-creditor-phone{display:block;color:var(--muted);font-size:11px;margin-top:3px}


.bill-items-header{display:flex;align-items:center;justify-content:space-between;gap:12px;margin:22px 0 10px}
.bill-items-table{border:1px solid var(--border);border-radius:14px;overflow:auto;background:#fff}
.bill-item-row{display:grid;grid-template-columns:34px 120px minmax(220px,1fr) 85px 120px 125px 42px;gap:8px;align-items:center;padding:10px;border-bottom:1px solid #edf0f5;min-width:760px}
.bill-item-row:last-child{border-bottom:0}
.bill-item-row select,.bill-item-row input{width:100%;padding:9px 8px;border:1px solid var(--border);border-radius:8px;background:#fff;min-width:0}
.bill-item-head{background:#f7f9fc;color:#667085;font-size:11px;font-weight:800;text-transform:uppercase}
.bill-line-no{font-weight:800;color:var(--muted);text-align:center}
.bill-line-total{font-size:13px;text-align:right}
.danger-icon{color:#b42318!important;background:#fff5f5!important}
.bill-summary-box{display:flex;justify-content:flex-end;gap:12px;margin-top:14px}
.bill-summary-box>div{min-width:150px;padding:13px 15px;border:1px solid var(--border);border-radius:12px;background:#fafbfd}
.bill-summary-box span,.bill-preview-summary span{display:block;color:var(--muted);font-size:11px;margin-bottom:5px}
.bill-summary-box strong{font-size:17px}
.bill-summary-box .bill-balance{background:#fff7ed;border-color:#fed7aa}
.bill-summary-box .bill-balance strong{color:#c2410c;font-size:20px}
.bill-multiple-badge{display:inline-block;margin-left:6px;padding:3px 6px;border-radius:999px;background:#eef4ff;color:#1262dc;font-size:9px;font-weight:800}
.bill-created-modal{max-width:620px}
.bill-preview-summary{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin-top:14px}
.bill-preview-summary>div{border:1px solid var(--border);border-radius:12px;padding:13px;background:#fafbfd}
.bill-preview-summary strong{font-size:16px}
@media(max-width:900px){.bill-summary-box{justify-content:stretch;flex-wrap:wrap}.bill-summary-box>div{flex:1;min-width:130px}}
@media(max-width:600px){.bill-preview-summary{grid-template-columns:1fr}.bill-items-header{align-items:flex-start;flex-direction:column}}


`;

export default function StyledApp() {
  return (
    <>
      <App />
      <style>{CSS}</style>
    </>
  );
}