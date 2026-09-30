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

  const businessId = staff?.business_id || null;

  const pages = {
    dashboard: <Dashboard businessId={businessId} go={go} refresh={refresh} />,
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
    sales: (
      <SalesPage
        businessId={businessId}
        staff={staff}
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
    attendance: <AttendancePage businessId={businessId} staff={staff} refresh={refresh} />,
    saasOverview: <SaaSAdminPage section="overview" refresh={refresh} />,
    saasCustomers: <SaaSAdminPage section="customers" refresh={refresh} />,
    saasSubscriptions: <SaaSAdminPage section="subscriptions" refresh={refresh} />,
    saasPayments: <SaaSAdminPage section="payments" refresh={refresh} />,
  };

  return (
    <div className="app-shell">
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

          <div className="nav-title">FEDHA</div>

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
            active={page === "attendance"}
            icon="🕐"
            text="Staff Attendance"
            onClick={() => go("attendance")}
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
                : page === "products"
                ? "Products"
                : page === "services"
                ? "Services"
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
                : page === "attendance"
                ? "Staff Attendance"
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

function Dashboard({ businessId, go, refresh }) {
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

    setLowProducts(low.slice(0, 8));
    setRecentSales(s.slice(0, 8));
    setLoading(false);
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
            <button className="text-btn" onClick={() => go("products")}>
              View all
            </button>
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
                  <div
                    className={
                      x.currentStock <= 0
                        ? "stock-danger"
                        : "stock-warning"
                    }
                  >
                    {x.currentStock} {x.unit}
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
      );
  }, [products, stockIn, sales, search]);

  async function addProduct(e) {
    e.preventDefault();
    setBusy(true);

    const { error } = await supabase.from("products").insert({
      business_id: businessId,
      product_name: form.product_name.trim(),
      unit: form.unit.trim() || "PCS",
      opening_qty: number(form.opening_qty),
      cost_per_each: number(form.cost_per_each),
      sell_per_each: number(form.sell_per_each),
      reorder_level: number(form.reorder_level),
      active: true,
    });

    if (error) {
      alert(error.message.includes("duplicate") ? "Bidhaa hiyo tayari ipo." : error.message);
    } else {
      setForm({
        product_name: "",
        unit: "PCS",
        opening_qty: "0",
        cost_per_each: "",
        sell_per_each: "",
        reorder_level: "5",
      });
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
        button="+ Add Product"
        onClick={() => setShowForm(!showForm)}
      />

      {showForm && (
        <div className="panel form-panel">
          <div className="panel-header">
            <div>
              <h3>Ongeza Bidhaa Mpya</h3>
              <span>Jaza taarifa za bidhaa</span>
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
                {busy ? "Inahifadhi..." : "Save Product"}
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
          <span className="record-count">{rows.length} products</span>
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
                    <button
                      className="danger-small"
                      onClick={() =>
                        deleteProduct(p.id, p.product_name)
                      }
                    >
                      Delete
                    </button>
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

function SalesPage({ businessId, staff, onChanged }) {
  const [products, setProducts] = useState([]);
  const [services, setServices] = useState([]);
  const [sales, setSales] = useState([]);
  const [stockIn, setStockIn] = useState([]);
  const [busy, setBusy] = useState(false);
  const [receipt, setReceipt] = useState(null);

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

    if (form.payment_method === "CREDIT" && !form.credit_customer_name.trim()) {
      alert("Weka jina la mteja anayechukua kwa mkopo.");
      return;
    }

    if (form.payment_method === "CREDIT") {
      const paidNow = number(form.credit_paid_amount);
      const saleTotal = qty * price;
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
    });

    setForm({
      ...form,
      quantity: "1",
      manual_name: "",
      sale_date: localDateTimeValue(),
    });

    await load();
    onChanged();
    setBusy(false);
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
                { value: "CREDIT", label: "MKOPO / CREDIT" },
              ]}
            />

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
              onClick={() => window.print()}
            >
              🖨 Print
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
          onClose={() => setReceipt(null)}
        />
      )}
    </div>
  );
}

function ReceiptModal({ receipt, onClose }) {
  function printReceipt() {
    window.print();
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
            <h2>Bless Stationery</h2>
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
            <small>Bless Stationery</small>
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

  async function addService(e) {
    e.preventDefault();

    const { error } = await supabase.from("services").insert({
      business_id: businessId,
      service_name: form.service_name.trim(),
      selling_price: number(form.selling_price),
      cost: number(form.cost),
      active: true,
    });

    if (error) {
      alert(error.message);
    } else {
      setForm({
        service_name: "",
        selling_price: "",
        cost: "0",
      });
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
        button="+ Add Service"
        onClick={() => setShow(!show)}
      />

      {show && (
        <div className="panel">
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
              <button className="primary-btn">Save Service</button>
            </div>
          </form>
        </div>
      )}

      <div className="service-grid">
        {services.map((s) => (
          <div className="service-card" key={s.id}>
            <div className="service-icon">⚙</div>
            <h3>{s.service_name}</h3>
            <span>Price</span>
            <strong>{money(s.selling_price)}</strong>
            <button
              className="danger-small"
              onClick={() => deleteService(s.id)}
            >
              Delete
            </button>
          </div>
        ))}

        {!services.length && (
          <EmptyState text="Hakuna services." />
        )}
      </div>
    </div>
  );
}

function ExpensesPage({ businessId, staff, onChanged }) {
  const [records, setRecords] = useState([]);
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
            label="Deposit Date & Time"
            type="datetime-local"
            value={form.deposit_date}
            onChange={(v) => setForm({ ...form, deposit_date: v })}
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
            <button className="primary-btn">Save Expense</button>
          </div>
        </form>
      </div>

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
            <button className="primary-btn">Save Deposit</button>
          </div>
        </form>
      </div>

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
      <tr><td>${formatDate(x.sale_date)}</td><td>${x.sale_type || "-"}</td><td>${x.quantity || 0}</td><td>${money(x.sales_total)}</td><td>${money(x.profit)}</td><td>${x.payment_method || "-"}</td></tr>
    `).join("");
    const expenseRows = report.expenses.map((x) => `
      <tr><td>${formatDate(x.expense_date)}</td><td>${x.expense_item || "-"}</td><td>${x.category || "-"}</td><td>${money(x.amount)}</td><td>${x.description || "-"}</td></tr>
    `).join("");
    const depositRows = report.deposits.map((x) => `
      <tr><td>${formatDate(x.deposit_date)}</td><td>${money(x.amount)}</td><td>${x.method || "-"}</td><td>${x.depositor || "-"}</td><td>${x.reference_no || "-"}</td></tr>
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
        <StatCard title="Staff Present" value={data.attendance.length} icon="👥" tone="green" />
        <StatCard title="Staff Absent" value={data.attendance.length} icon="◷" tone="purple" />
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
  const [form, setForm] = useState({
    customer_name: "",
    customer_phone: "",
    product_id: "",
    service_id: "",
    item_type: "PRODUCT",
    amount: "",
    paid_amount: "0",
    due_date: "",
    note: "",
  });
  const [paying, setPaying] = useState(null);
  const [paymentAmount, setPaymentAmount] = useState("");
  const [search, setSearch] = useState("");
  const [busy, setBusy] = useState(false);

  async function load() {
    if (!businessId) return;
    const [{ data: c, error }, { data: p }, { data: sv }] = await Promise.all([
      supabase.from("credit_transactions").select("*, sales(sale_date, sale_type, quantity, selling_price, products(product_name), services(service_name))").eq("business_id", businessId).order("due_date", { ascending: true }),
      supabase.from("products").select("*").eq("business_id", businessId).eq("active", true).order("product_name"),
      supabase.from("services").select("*").eq("business_id", businessId).eq("active", true).order("service_name"),
    ]);
    if (error) { alert(error.message); return; }
    setCredits(c || []); setProducts(p || []); setServices(sv || []);
  }
  useEffect(() => { load(); }, [businessId, refresh]);

  async function addDebt(e) {
    e.preventDefault();
    const original = number(form.amount);
    const paid = number(form.paid_amount);
    if (!form.customer_name.trim() || original <= 0) { alert("Weka jina la mteja na kiasi sahihi."); return; }
    if (paid < 0 || paid > original) { alert("Kiasi alicholipa hakiwezi kuzidi deni."); return; }
    if (!form.due_date) { alert("Chagua tarehe ya mwisho ya kulipa."); return; }
    setBusy(true);
    const balance = Math.max(0, original - paid);
    const { error } = await supabase.from("credit_transactions").insert({
      business_id: businessId,
      customer_name: form.customer_name.trim(),
      customer_phone: form.customer_phone.trim() || null,
      product_id: form.item_type === "PRODUCT" ? (form.product_id || null) : null,
      service_id: form.item_type === "SERVICE" ? (form.service_id || null) : null,
      original_amount: original,
      paid_amount: paid,
      balance,
      status: balance <= 0 ? "PAID" : paid > 0 ? "PARTIAL" : "UNPAID",
      due_date: form.due_date,
      note: form.note.trim() || null,
      created_by: staff?.id || null,
    });
    if (error) alert(error.message);
    else {
      setForm({ customer_name:"", customer_phone:"", product_id:"", service_id:"", item_type:"PRODUCT", amount:"", paid_amount:"0", due_date:"", note:"" });
      await load();
    }
    setBusy(false);
  }

  async function payDebt(row) {
    const pay = number(paymentAmount);
    if (pay <= 0 || pay > number(row.balance)) { alert("Kiasi cha malipo si sahihi."); return; }
    const paid = number(row.paid_amount) + pay;
    const balance = Math.max(0, number(row.original_amount) - paid);
    const { error } = await supabase.from("credit_transactions").update({
      paid_amount: paid, balance, status: balance <= 0 ? "PAID" : "PARTIAL", last_payment_at: new Date().toISOString()
    }).eq("id", row.id).eq("business_id", businessId);
    if (error) alert(error.message); else { setPaying(null); setPaymentAmount(""); await load(); }
  }

  function itemName(x) {
    return x.sales?.products?.product_name || x.sales?.services?.service_name || products.find(p=>p.id===x.product_id)?.product_name || services.find(s=>s.id===x.service_id)?.service_name || "-";
  }
  function sendReminder(row) {
    if (!row.customer_phone) { alert("Mteja hana namba ya simu."); return; }
    const phone = row.customer_phone.replace(/[^0-9]/g, "");
    const msg = `Habari ${row.customer_name}, tunakukumbusha kuwa una deni la TZS ${money(row.balance)} katika Bless Stationery. Tarehe ya mwisho ya malipo ni ${row.due_date || "leo"}. Tafadhali lipa kwa wakati. Asante.`;
    window.open(`https://wa.me/${phone}?text=${encodeURIComponent(msg)}`, "_blank");
  }

  const filtered = credits.filter(x => `${x.customer_name} ${x.customer_phone || ""}`.toLowerCase().includes(search.toLowerCase()));
  const totalDebt = filtered.reduce((a,x)=>a+number(x.balance),0);
  const totalOriginal = filtered.reduce((a,x)=>a+number(x.original_amount),0);
  const totalPaid = filtered.reduce((a,x)=>a+number(x.paid_amount),0);

  return <div>
    <PageTitle title="Madeni / Credit" subtitle="Simamia bidhaa na huduma zilizochukuliwa kwa mkopo." />
    <div className="stats-grid">
      <StatCard title="Madeni Yote" value={money(totalDebt)} icon="💳" tone="red" />
      <StatCard title="Jumla ya Mikopo" value={money(totalOriginal)} icon="📒" tone="blue" />
      <StatCard title="Yaliyolipwa" value={money(totalPaid)} icon="✓" tone="green" />
      <StatCard title="Wadaiwa" value={filtered.filter(x=>number(x.balance)>0).length} icon="👥" tone="orange" />
    </div>
    <div className="panel">
      <div className="panel-header"><div><h3>Ongeza Deni</h3><span>Weka mteja, product/service, kiasi alicholipa na tarehe ya mwisho.</span></div></div>
      <form onSubmit={addDebt} className="form-grid">
        <Field label="Jina la Mteja" value={form.customer_name} onChange={v=>setForm({...form,customer_name:v})} required />
        <Field label="Simu ya Mteja" value={form.customer_phone} onChange={v=>setForm({...form,customer_phone:v})} placeholder="2557XXXXXXXX" />
        <SelectField label="Aina" value={form.item_type} onChange={v=>setForm({...form,item_type:v,product_id:"",service_id:""})} options={[{value:"PRODUCT",label:"PRODUCT"},{value:"SERVICE",label:"SERVICE"}]} />
        {form.item_type === "PRODUCT" ? <SelectField label="Product" value={form.product_id} onChange={v=>setForm({...form,product_id:v})} options={products.map(p=>({value:p.id,label:p.product_name}))} /> : <SelectField label="Service" value={form.service_id} onChange={v=>setForm({...form,service_id:v})} options={services.map(s=>({value:s.id,label:s.service_name}))} />}
        <Field label="Kiasi Jumla" type="number" value={form.amount} onChange={v=>setForm({...form,amount:v})} required />
        <Field label="Kiasi Alicholipa" type="number" value={form.paid_amount} onChange={v=>setForm({...form,paid_amount:v})} />
        <Field label="Deni Litalipwa Tarehe" type="date" value={form.due_date} onChange={v=>setForm({...form,due_date:v})} required />
        <Field label="Maelezo" value={form.note} onChange={v=>setForm({...form,note:v})} />
        <div className="notice full">Salio: <strong>{money(Math.max(0, number(form.amount)-number(form.paid_amount)))}</strong></div>
        <div className="form-actions full"><button className="primary-btn" disabled={busy}>{busy ? "Inahifadhi..." : "+ Weka Deni"}</button></div>
      </form>
    </div>
    <div className="panel">
      <div className="toolbar"><input className="search-input" value={search} onChange={e=>setSearch(e.target.value)} placeholder="Tafuta mteja..." /></div>
      <div className="table-wrap"><table><thead><tr><th>Mteja</th><th>Product/Service</th><th>Deni</th><th>Amelipa</th><th>Salio</th><th>Due Date</th><th>Status</th><th>Action</th></tr></thead><tbody>
        {filtered.map(x=><tr key={x.id}><td><strong>{x.customer_name}</strong><br/><small>{x.customer_phone || "-"}</small></td><td>{itemName(x)}</td><td>{money(x.original_amount)}</td><td>{money(x.paid_amount)}</td><td><strong>{money(x.balance)}</strong></td><td>{x.due_date || "-"}</td><td>{x.status}</td><td>{number(x.balance)>0 && <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{paying===x.id ? <><input type="number" value={paymentAmount} onChange={e=>setPaymentAmount(e.target.value)} placeholder="Kiasi" /><button className="primary-btn" onClick={()=>payDebt(x)}>Lipa</button><button className="secondary-btn" onClick={()=>setPaying(null)}>X</button></> : <button className="secondary-btn" onClick={()=>setPaying(x.id)}>+ Malipo</button>}<button className="secondary-btn" onClick={()=>sendReminder(x)}>📲 Kumbusha</button></div>}</td></tr>)}
        {!filtered.length && <tr><td colSpan="8">Hakuna madeni yaliyopatikana.</td></tr>}
      </tbody></table></div>
    </div>
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
    <div className="stats-grid"><StatCard title="Kuingia" value="08:30" icon="🌅" tone="blue" /><StatCard title="Kutoka" value="21:30" icon="🌙" tone="purple" /><StatCard title="Present" value={attendance.length} icon="✓" tone="green" /><StatCard title="Absent" value={attendance.length} icon="◷" tone="red" /></div>
    {staff?.active && <div className="panel"><div className="panel-header"><div><h3>{staff.staff_name} — Leo</h3><span>System time ndiyo unaotumika; staff haandiki muda.</span></div><div style={{display:"flex",gap:8}}><button className="primary-btn" disabled={busy||!!todayRecord} onClick={checkIn}>✓ SIGN IN</button><button className="secondary-btn" disabled={busy||!todayRecord||!!todayRecord.check_out_at} onClick={checkOut}>↪ SIGN OUT</button></div></div>{todayRecord&&<div className="notice">Kuingia: <strong>{localTime(todayRecord.check_in_at)}</strong> · Kutoka: <strong>{localTime(todayRecord.check_out_at)}</strong> · Late: <strong>{minutesLate(todayRecord.check_in_at)} min</strong></div>}</div>}
    <div className="panel"><div className="toolbar"><input type="date" value={date} onChange={e=>setDate(e.target.value)} /></div><div className="table-wrap"><table><thead><tr><th>Staff</th><th>Kuingia</th><th>Kutoka</th><th>Late</th><th>Working Hours</th><th>Status</th></tr></thead><tbody>{staffRows.map(s=>{const a=attendance.find(x=>x.staff_id===s.id);return <tr key={s.id}><td><strong>{s.staff_name}</strong></td><td>{localTime(a?.check_in_at)}</td><td>{localTime(a?.check_out_at)}</td><td>{a?`${minutesLate(a.check_in_at)} min`:"-"}</td><td>{worked(a?.check_in_at,a?.check_out_at)}</td><td>{a?(a.check_out_at?"PRESENT":"IN WORK"):"ABSENT"}</td></tr>})}</tbody></table></div></div>
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

const CSS = `
* {
  box-sizing: border-box;
}

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
  background: white;
  color: #1262dc;
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
  background: #071b38;
  color: white;
  min-height: 100vh;
  padding: 20px 14px;
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
`;

export default function StyledApp() {
  return (
    <>
      <App />
      <style>{CSS}</style>
    </>
  );
}