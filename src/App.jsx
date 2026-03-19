import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

const NG = "#00C73C";
const STORAGE_KEY = "forcool_v4";
const n = (v) => parseFloat(String(v || "").replace(/,/g, "")) || 0;
const fmt = (v) => Math.round(v).toLocaleString();
const cpc = (cost, clicks) => n(clicks) > 0 ? Math.round(n(cost) / n(clicks)) : 0;
const roasVal = (rev, cost) => n(cost) > 0 ? Math.round(n(rev) / n(cost) * 100) : 0;
const kwCount = (str) => str ? str.split("\n").filter(k => k.trim()).length : 0;

// 제외 키워드 자동 교차 계산 (쌍 구조)
// 0(가격비교 메인) ↔ 1(가격비교 서브)
// 2(N+스토어 메인) ↔ 3(N+스토어 서브)
// 같은 광고유형 내에서만 메인/서브 교차 차단
const EXCL_PAIRS = [[0, 1], [2, 3]];
const calcExcludeKeywords = (groups) => {
  return groups.map((g, i) => {
    const pair = EXCL_PAIRS.find(p => p.includes(i));
    if (!pair) return g;
    const partnerId = pair.find(j => j !== i);
    const partner = groups[partnerId];
    const myKws = new Set((g.registerKeywords || "").split("\n").map(k => k.trim()).filter(Boolean));
    const excludeKws = (partner?.registerKeywords || "")
      .split("\n")
      .map(k => k.trim())
      .filter(k => k && !myKws.has(k)); // 내 등록 키워드와 겹치는 건 제외에서 제거
    return { ...g, excludeKeywords: [...new Set(excludeKws)].join("\n") };
  });
};

const roasBadge = (r) => {
  if (r >= 5000) return { bg: "#D1FAE5", color: "#065F46", label: "입찰가 인상 권장 ▲" };
  if (r >= 1000) return { bg: "#FEF3C7", color: "#92400E", label: "유지 ─" };
  if (r > 0) return { bg: "#FEE2E2", color: "#991B1B", label: "OFF 검토 ▼" };
  return null;
};

const GRP_META = [
  { label: "1번", sub: "가격비교 메인", color: "#00873A", bg: "#F0FDF4", border: "#D1FAE5" },
  { label: "2번", sub: "가격비교 서브", color: "#7C3AED", bg: "#F5F3FF", border: "#EDE9FE" },
  { label: "3번", sub: "N+스토어 메인", color: "#B45309", bg: "#FFFBEB", border: "#FDE68A" },
  { label: "4번", sub: "N+스토어 서브", color: "#0369A1", bg: "#F0F9FF", border: "#BAE6FD" },
];

const mkGroup = (i) => ({
  id: i,
  productName: "", productNames: [], registerKeywords: "", excludeKeywords: "",
  bidPrice: "", bidPc: [], bidMo: [], dailyBudget: "",
  groupLabel: "", groupKeywords: "",
  impressions: "", clicks: "", cost: "", conversions: "", revenue: "",
});

const mkCamp = (num = "", name = "", meta = {}) => ({
  id: `c${Date.now()}${Math.random().toString(36).slice(2, 5)}`,
  number: num, name,
  adType: meta.adType || "쇼검",
  brand: meta.brand || "스타리온",
  model: meta.model || "",
  mainKw: meta.mainKw || "",
  campTotalBudget: meta.campTotalBudget || "",
  status: "세팅중", // 세팅중 | 진행중 | 일시정지 | 종료
  history: [{ date: new Date().toISOString(), action: "캠페인 생성" }],
  groups: [0, 1, 2, 3].map(mkGroup),
});

const buildCampName = (adType, brand, model, mainKw) => {
  const parts = [model, mainKw].filter(Boolean);
  return `${adType}#${brand}${parts.length ? "/" + parts.join("/") : ""}`;
};

const S = {
  input: {
    width: "100%", border: "1.5px solid #E5E7EB", borderRadius: 8,
    padding: "8px 12px", fontSize: 13, outline: "none",
    boxSizing: "border-box", background: "#fff", color: "#111", fontFamily: "inherit",
  },
  label: {
    fontSize: 11, fontWeight: 600, color: "#333",
    display: "block", marginBottom: 5, letterSpacing: "0.04em",
  },
  btn: (active, color = NG) => ({
    padding: "7px 14px", borderRadius: 7, cursor: "pointer", fontSize: 12,
    fontWeight: 600, border: `1.5px solid ${active ? color : "#E5E7EB"}`,
    background: active ? `${color}18` : "#fff",
    color: active ? color : "#6B7280", transition: "all 0.15s",
  }),
  card: {
    background: "#fff", borderRadius: 12, border: "1.5px solid #F0F0F0",
    padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.04)",
  },
};

// ── XLSX Parser ───────────────────────────────────────────────────────────
const parseXlsx = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const result = {};
      wb.SheetNames.forEach(name => {
        const ws = wb.Sheets[name];
        result[name] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      });
      resolve({ name: file.name, sheets: result });
    } catch (err) { reject(err); }
  };
  reader.readAsArrayBuffer(file);
});

// ── Claude API Call with Image ───────────────────────────────────────────
const getApiKey = () => localStorage.getItem("forcool_api_key") || "";
const saveApiKey = (key) => localStorage.setItem("forcool_api_key", key);

const callClaudeWithImage = async (prompt, systemPrompt, imageBase64) => {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error("API_KEY_MISSING");
  const content = [];
  if (imageBase64) {
    const mimeMatch = imageBase64.match(/data:([^;]+);base64,/);
    const mediaType = mimeMatch ? mimeMatch[1] : "image/jpeg";
    const base64Data = imageBase64.split(",")[1];
    content.push({ type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } });
    content.push({ type: "text", text: "위 이미지는 광고할 상품의 대표 이미지입니다. 이미지를 참고해서 상품 카테고리와 특성을 정확히 파악한 후 광고 세팅을 해주세요." });
  }
  content.push({ type: "text", text: prompt });
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 3000,
      system: systemPrompt,
      messages: [{ role: "user", content }],
    }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.content?.[0]?.text || "";
};

// ── Toast ─────────────────────────────────────────────────────────────────
function Toast({ msg }) {
  if (!msg) return null;
  return (
    <div style={{ position: "fixed", bottom: 28, left: "50%", transform: "translateX(-50%)", background: "#111", color: "#fff", padding: "10px 22px", borderRadius: 24, fontSize: 13, zIndex: 9999, boxShadow: "0 4px 20px rgba(0,0,0,0.3)", fontWeight: 500 }}>
      {msg}
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────
export default function App() {
  const [camps, setCamps] = useState([]);
  const [selId, setSelId] = useState(null);
  const [campModal, setCampModal] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [toast, setToast] = useState("");
  const [showApiModal, setShowApiModal] = useState(false);
  const [hasApiKey, setHasApiKey] = useState(() => !!localStorage.getItem("forcool_api_key"));

  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get(STORAGE_KEY);
        if (r?.value) {
          const d = JSON.parse(r.value);
          if (d.camps?.length) { setCamps(d.camps); setSelId(d.selId || d.camps[0].id); }
        }
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try { await window.storage.set(STORAGE_KEY, JSON.stringify({ camps, selId })); }
      catch (e) {}
    })();
  }, [camps, selId, loaded]);

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(""), 2500); };

  const camp = camps.find(c => c.id === selId);
  const totalBudget = camp ? camp.groups.reduce((s, g) => s + n(g.dailyBudget), 0) : 0;

  const updateGroup = (cid, gid, field, val) => {
    setCamps(prev => prev.map(c => {
      if (c.id !== cid) return c;
      const updated = c.groups.map(g => g.id !== gid ? g : { ...g, [field]: val });
      // 등록 키워드 변경 시 제외 키워드 자동 재계산
      const newGroups = field === "registerKeywords" ? calcExcludeKeywords(updated) : updated;
      return { ...c, groups: newGroups };
    }));
  };

  const toggleStatus = (id, e) => {
    e.stopPropagation();
    setCamps(prev => prev.map(c => {
      if (c.id !== id) return c;
      const next = c.status === "진행중" ? "세팅중" : "진행중";
      const entry = { date: new Date().toISOString(), action: next === "진행중" ? "광고 ON" : "광고 OFF" };
      return { ...c, status: next, history: [...(c.history || []), entry] };
    }));
  };

  const saveCampMeta = (meta) => {
    const name = buildCampName(meta.adType, meta.brand, meta.model, meta.mainKw);
    if (campModal === "new") {
      const c = mkCamp(meta.number, name, meta);
      setCamps(prev => [...prev, c]);
      setSelId(c.id);
      showToast(`${meta.number} 캠페인 추가됨`);
    } else {
      const entry = { date: new Date().toISOString(), action: "캠페인 정보 수정" };
      setCamps(prev => prev.map(c => c.id !== campModal ? c : {
        ...c, ...meta, name, history: [...(c.history || []), entry]
      }));
      showToast("저장됨");
    }
    setCampModal(null);
  };

  const delCamp = (id, e) => {
    e.stopPropagation();
    const remaining = camps.filter(c => c.id !== id);
    setCamps(remaining);
    if (selId === id) setSelId(remaining[0]?.id || null);
    showToast("캠페인 삭제됨");
  };

  const copyCamp = (id, e) => {
    e.stopPropagation();
    const src = camps.find(c => c.id === id);
    if (!src) return;
    const newId = `c${Date.now()}${Math.random().toString(36).slice(2, 5)}`;
    const copy = {
      ...JSON.parse(JSON.stringify(src)),
      id: newId,
      number: src.number + "-복사",
      status: "세팅중",
      history: [{ date: new Date().toISOString(), action: `'${src.number}' 복사 생성` }],
      memo: "",
    };
    setCamps(prev => [...prev, copy]);
    setSelId(newId);
    showToast("캠페인 복사됨!");
  };

  const updateMemo = (id, val) => {
    setCamps(prev => prev.map(c => c.id !== id ? c : { ...c, memo: val }));
  };

  const applyGeneratedGroups = (groups, summary) => {
    if (!camp) return;
    const limited = groups.slice(0, 4);
    const mainKwStr = (summary?.mainKeywords || []).map(k => typeof k === "object" ? k.kw : k).join("\n");
    const subKwStr = (summary?.subKeywords || []).map(k => typeof k === "object" ? k.kw : k).join("\n");
    const withMeta = limited.map((g, i) => {
      const kwList = i % 2 === 0 ? (summary?.mainKeywords || []) : (summary?.subKeywords || []);
      const kwStr = JSON.stringify(kwList.map(k => typeof k === "object" ? k : { kw: k, vol: null }));
      const regKws = i % 2 === 0 ? mainKwStr : subKwStr;
      return { ...g, groupLabel: i % 2 === 0 ? "메인" : "서브", groupKeywords: kwStr, registerKeywords: regKws };
    });
    withMeta[2] = { ...withMeta[2], registerKeywords: mainKwStr, groupKeywords: withMeta[0].groupKeywords };
    withMeta[3] = { ...withMeta[3], registerKeywords: subKwStr, groupKeywords: withMeta[1].groupKeywords };
    const withExclude = calcExcludeKeywords(withMeta);
    const entry = { date: new Date().toISOString(), action: "AI 세팅 적용" };
    setCamps(prev => prev.map(c => {
      if (c.id !== camp.id) return c;
      const newGroups = c.groups.map((g, i) => ({ ...g, ...withExclude[i], id: g.id }));
      return { ...c, groups: newGroups, history: [...(c.history || []), entry] };
    }));
    showToast("✨ 4개 그룹 자동 생성 완료!");
  };

  return (
    <div style={{ fontFamily: "'Pretendard', 'Apple SD Gothic Neo', -apple-system, sans-serif", background: "#F4F5F7", minHeight: "100vh", display: "flex", flexDirection: "column" }}>
      {/* Header */}
      <header style={{ background: "#fff", borderBottom: "1.5px solid #F0F0F0", padding: "0 20px", height: 54, display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 100, flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 28, height: 28, borderRadius: 7, background: NG, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <span style={{ color: "#fff", fontSize: 14, fontWeight: 700 }}>N</span>
          </div>
          <span style={{ fontWeight: 700, fontSize: 15, color: "#111" }}>AI 쇼핑검색광고 세팅 v3.7</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 11, color: "#444" }}>캠페인 {camps.length}개 운영 중</span>
          <button onClick={() => setShowApiModal(true)} title={hasApiKey ? "API 연결됨" : "API 키 설정 필요"}
            style={{ width: 28, height: 28, borderRadius: 7, border: hasApiKey ? "0.5px solid #D1FAE5" : "0.5px solid #FCA5A5", background: hasApiKey ? "#F0FDF4" : "#FFF5F5", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14 }}>
            {hasApiKey ? "🔑" : "⚠️"}
          </button>
        </div>
      </header>

      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        {/* Sidebar */}
        <aside style={{ width: 220, background: "#fff", borderRight: "1.5px solid #F0F0F0", display: "flex", flexDirection: "column", overflow: "hidden" }}>
          <div style={{ padding: "12px 12px 8px", display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
            <span style={{ fontSize: 10, fontWeight: 700, color: "#444", letterSpacing: "0.08em" }}>캠페인</span>
            <button onClick={() => setCampModal("new")} style={{ background: NG, color: "#fff", border: "none", borderRadius: 6, padding: "4px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>+ 추가</button>
          </div>
          <div style={{ flex: 1, overflowY: "auto" }}>
            {camps.length === 0 && (
              <div style={{ padding: "20px 14px", color: "#666", fontSize: 12, textAlign: "center" }}>
                <div style={{ fontSize: 24, marginBottom: 8 }}>📋</div>
                캠페인을 추가하세요
              </div>
            )}
            {camps.map(c => {
              const sel = selId === c.id;
              const campBudget = c.groups.reduce((s, g) => s + n(g.dailyBudget), 0);
              const isOn = c.status === "진행중";
              const statusColor = { "진행중": "#00C73C", "세팅중": "#F59E0B", "일시정지": "#9CA3AF", "종료": "#EF4444" }[c.status || "세팅중"];
              return (
                <div key={c.id} onClick={() => setSelId(c.id)}
                  style={{ padding: "10px 12px", cursor: "pointer", borderLeft: `3px solid ${sel ? NG : "transparent"}`, background: sel ? "#F0FDF4" : "transparent", borderBottom: "1px solid #F9FAFB" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 6 }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 5, marginBottom: 2 }}>
                        <span style={{ fontSize: 10, color: "#444", fontWeight: 600 }}>{c.number}</span>
                        <span style={{ fontSize: 9, fontWeight: 700, padding: "1px 5px", borderRadius: 8, background: `${statusColor}20`, color: statusColor }}>{c.status || "세팅중"}</span>
                      </div>
                      <div style={{ fontSize: 12, fontWeight: sel ? 700 : 500, color: sel ? "#00873A" : "#111", lineHeight: 1.4, wordBreak: "break-all" }}>{c.name}</div>
                      {campBudget > 0 && <div style={{ fontSize: 10, color: "#444", marginTop: 2 }}>일 {fmt(campBudget)}원</div>}
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 4, flexShrink: 0 }}>
                      <button onClick={(e) => toggleStatus(c.id, e)}
                        title={isOn ? "OFF로 변경" : "ON으로 변경"}
                        style={{ width: 36, height: 20, borderRadius: 10, border: "none", background: isOn ? NG : "#D1D5DB", cursor: "pointer", position: "relative", flexShrink: 0, transition: "background 0.2s" }}>
                        <span style={{ position: "absolute", top: 2, left: isOn ? 18 : 2, width: 16, height: 16, borderRadius: "50%", background: "#fff", transition: "left 0.2s", boxShadow: "0 1px 3px rgba(0,0,0,0.2)" }} />
                      </button>
                      <div style={{ display: "flex", gap: 3 }}>
                        <button onClick={(e) => copyCamp(c.id, e)} title="캠페인 복사" style={{ background: "none", border: "none", color: "#555", cursor: "pointer", fontSize: 11, padding: "0" }}>⎘</button>
                        <button onClick={(e) => delCamp(c.id, e)} style={{ background: "none", border: "none", color: "#666", cursor: "pointer", fontSize: 12, padding: "0" }}>✕</button>
                      </div>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </aside>

        {/* Main */}
        <main style={{ flex: 1, overflow: "auto", padding: "20px 20px 40px" }}>
          {camp ? (
            <>
              {/* Campaign Header */}
              <CampHeader camp={camp} onEdit={() => setCampModal(camp.id)} totalBudget={totalBudget} showToast={showToast} toggleStatus={(e) => toggleStatus(camp.id, e)} />

              <HistoryPanel camp={camp} />
              <MemoPanel camp={camp} onUpdate={(val) => updateMemo(camp.id, val)} />

              {/* AI Generate Panel */}
              <AIGeneratePanel key={camp.id} camp={camp} onApply={applyGeneratedGroups} showToast={showToast} />

              {/* 4 Group Grid - 분석 후에만 표시 */}
              {camp.groups.some(g => g.productName) && (
                <FourGroupGrid camp={camp} onChange={updateGroup} showToast={showToast} />
              )}
            </>
          ) : (
            <div style={{ textAlign: "center", padding: "80px 20px", color: "#444" }}>
              <div style={{ fontSize: 48, marginBottom: 16 }}>📦</div>
              <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>캠페인을 추가해주세요</div>
              <div style={{ fontSize: 13 }}>왼쪽 + 추가 버튼을 눌러 캠페인을 만들어보세요</div>
              <button onClick={() => setCampModal("new")} style={{ marginTop: 20, background: NG, color: "#fff", border: "none", borderRadius: 9, padding: "11px 28px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>+ 새 캠페인 추가</button>
            </div>
          )}
        </main>
      </div>

      {showApiModal && (
        <ApiKeyModal
          onSave={(key) => { saveApiKey(key); setHasApiKey(true); setShowApiModal(false); showToast("API 키 저장됨!"); }}
          onClose={() => setShowApiModal(false)}
        />
      )}
      {campModal !== null && (
        <CampModal
          camp={campModal !== "new" ? camps.find(c => c.id === campModal) : null}
          onSave={saveCampMeta}
          onClose={() => setCampModal(null)}
        />
      )}
      <Toast msg={toast} />
    </div>
  );
}

// ── Camp Header ───────────────────────────────────────────────────────────
function CampHeader({ camp, onEdit, totalBudget, showToast, toggleStatus }) {
  const isOn = camp.status === "진행중";
  const statusColor = { "진행중": "#00C73C", "세팅중": "#F59E0B", "일시정지": "#9CA3AF", "종료": "#EF4444" }[camp.status || "세팅중"];
  const [copied, setCopied] = useState(false);
  const copy = () => {
    navigator.clipboard.writeText(camp.name);
    setCopied(true); setTimeout(() => setCopied(false), 2000);
    showToast("캠페인명 복사됨!");
  };
  return (
    <div style={{ ...S.card, marginBottom: 14 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
        <div style={{ flex: 1 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
            <span style={{ fontSize: 10, color: "#444", fontWeight: 700, letterSpacing: "0.06em" }}>캠페인 {camp.number}</span>
            <span style={{ fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 10, background: `${statusColor}20`, color: statusColor }}>{camp.status || "세팅중"}</span>
            <button onClick={toggleStatus}
              style={{ width: 40, height: 22, borderRadius: 11, border: "none", background: isOn ? NG : "#D1D5DB", cursor: "pointer", position: "relative", transition: "background 0.2s" }}>
              <span style={{ position: "absolute", top: 3, left: isOn ? 20 : 3, width: 16, height: 16, borderRadius: "50%", background: "#fff", transition: "left 0.2s", boxShadow: "0 1px 3px rgba(0,0,0,0.2)" }} />
            </button>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
            <span style={{ fontSize: 17, fontWeight: 700, color: "#111" }}>{camp.name}</span>
            <button onClick={copy} style={{ background: copied ? "#E1F5EE" : "#F9FAFB", border: `1.5px solid ${copied ? NG : "#E5E7EB"}`, borderRadius: 6, padding: "3px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700, color: copied ? NG : "#111" }}>
              {copied ? "✓ 복사됨" : "📋 복사"}
            </button>
            <button onClick={onEdit} style={{ background: "#F9FAFB", border: "1.5px solid #E5E7EB", borderRadius: 6, padding: "3px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700, color: "#111" }}>✏️ 편집</button>
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {[
              { label: "광고유형", val: camp.adType, c: "#2563EB", bg: "#EFF6FF" },
              { label: "브랜드", val: camp.brand, c: "#7C3AED", bg: "#F5F3FF" },
              { label: "모델명", val: camp.model, c: "#374151", bg: "#F3F4F6" },
              { label: "메인키워드", val: camp.mainKw, c: "#065F46", bg: "#F0FDF4" },
            ].filter(b => b.val).map(b => (
              <span key={b.label} style={{ fontSize: 11, fontWeight: 600, padding: "3px 9px", borderRadius: 20, background: b.bg, color: b.c }}>{b.label}: {b.val}</span>
            ))}
          </div>
        </div>
        <div style={{ textAlign: "right", flexShrink: 0 }}>
          <div style={{ fontSize: 10, color: "#444", marginBottom: 2 }}>일 총 예산</div>
          <div style={{ fontSize: 22, fontWeight: 800, color: NG }}>{camp.campTotalBudget ? fmt(n(camp.campTotalBudget)) : fmt(totalBudget)}원</div>
        </div>
      </div>
    </div>
  );
}

// ── History Panel ─────────────────────────────────────────────────────────
function HistoryPanel({ camp }) {
  const [open, setOpen] = useState(false);
  const history = [...(camp.history || [])].reverse();
  const fmtDate = (iso) => {
    const d = new Date(iso);
    return `${d.getFullYear().toString().slice(2)}.${String(d.getMonth()+1).padStart(2,"0")}.${String(d.getDate()).padStart(2,"0")} ${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
  };
  const actionColor = (a) => {
    if (a.includes("ON")) return "#00C73C";
    if (a.includes("OFF")) return "#9CA3AF";
    if (a.includes("AI")) return "#7C3AED";
    if (a.includes("생성")) return "#2563EB";
    return "#374151";
  };
  return (
    <div style={{ ...S.card, marginBottom: 14 }}>
      <button onClick={() => setOpen(!open)}
        style={{ background: "none", border: "none", cursor: "pointer", padding: 0, display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
        <span style={{ fontSize: 12, fontWeight: 700, color: "#111" }}>📋 활동 기록 <span style={{ fontWeight: 500, color: "#555", fontSize: 11 }}>({history.length}건)</span></span>
        <span style={{ fontSize: 11, color: "#555" }}>{open ? "▲ 접기" : "▼ 펼치기"}</span>
      </button>
      {open && (
        <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 4, maxHeight: 180, overflowY: "auto" }}>
          {history.length === 0 && <div style={{ fontSize: 12, color: "#555", textAlign: "center", padding: "12px 0" }}>기록 없음</div>}
          {history.map((h, i) => (
            <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 8px", background: "#F9FAFB", borderRadius: 6, borderLeft: `3px solid ${actionColor(h.action)}` }}>
              <span style={{ fontSize: 12, fontWeight: 600, color: actionColor(h.action) }}>{h.action}</span>
              <span style={{ fontSize: 11, color: "#555" }}>{fmtDate(h.date)}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Memo Panel ────────────────────────────────────────────────────────────
function MemoPanel({ camp, onUpdate }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ ...S.card, marginBottom: 14 }}>
      <button onClick={() => setOpen(!open)}
        style={{ background: "none", border: "none", cursor: "pointer", padding: 0, display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
        <span style={{ fontSize: 12, fontWeight: 700, color: "#111" }}>📝 메모 {camp.memo ? <span style={{ fontWeight: 400, color: "#555", fontSize: 11 }}>(작성됨)</span> : <span style={{ fontWeight: 400, color: "#aaa", fontSize: 11 }}>(없음)</span>}</span>
        <span style={{ fontSize: 11, color: "#555" }}>{open ? "▲ 접기" : "▼ 펼치기"}</span>
      </button>
      {open && (
        <textarea
          value={camp.memo || ""}
          onChange={e => onUpdate(e.target.value)}
          placeholder="인수인계 내용, 특이사항, 전략 메모 등 자유롭게 작성하세요"
          style={{ ...S.input, marginTop: 10, height: 100, resize: "vertical", fontSize: 12, lineHeight: 1.7, color: "#111" }}
        />
      )}
    </div>
  );
}
// 단계: input → analyzing → review → applying
function AIGeneratePanel({ camp, onApply, showToast }) {
  const [open, setOpen] = useState(true);
  const [step, setStep] = useState("input"); // input | analyzing | review | done
  const [productName, setProductName] = useState("");
  const [files, setFiles] = useState({ relKw: null, adGroup: null, keywords: null, image: null });
  const [imagePreview, setImagePreview] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("");
  const [analysis, setAnalysis] = useState(null); // { summary, groups, rawData }
  const [comment, setComment] = useState("");

  const fileLabels = [
    { key: "image", label: "대표 이미지", accept: "image/*", icon: "🖼️", color: "#DC2626", bg: "#FFF5F5", border: "#FCA5A5", isImage: true },
    { key: "relKw", label: "연관 검색어", accept: ".xlsx", icon: "🔍", color: "#7C3AED", bg: "#F5F3FF", border: "#EDE9FE" },
    { key: "adGroup", label: "광고 리포트", accept: ".xlsx", icon: "📋", color: "#0369A1", bg: "#F0F9FF", border: "#BAE6FD" },
    { key: "keywords", label: "노출 키워드 리스트", accept: ".xlsx", icon: "🔑", color: "#B45309", bg: "#FFFBEB", border: "#FDE68A" },
  ];

  const handleFile = (key, e) => {
    const file = e.target.files?.[0];
    if (file) {
      setFiles(prev => ({ ...prev, [key]: file }));
      if (key === "image") {
        const reader = new FileReader();
        reader.onload = (ev) => setImagePreview(ev.target.result);
        reader.readAsDataURL(file);
      }
    }
    e.target.value = "";
  };
  const removeFile = (key) => {
    setFiles(prev => ({ ...prev, [key]: null }));
    if (key === "image") setImagePreview(null);
  };

  const buildDataStr = async () => {
    const parsed = {};
    for (const [key, file] of Object.entries(files)) {
      if (file && key !== "image") parsed[key] = await parseXlsx(file);
    }
    return Object.entries(parsed).map(([k, v]) => {
      const label = { relKw: "연관 검색어 데이터", adGroup: "광고 리포트 데이터", keywords: "노출 키워드 리스트" }[k];
      const sheets = Object.entries(v.sheets).map(([name, rows]) =>
        `[시트: ${name}]\n${rows.slice(0, 80).map(r => r.join("\t")).join("\n")}`
      ).join("\n\n");
      return `=== ${label} (${v.name}) ===\n${sheets}`;
    }).join("\n\n");
  };

  const systemPrompt = `당신은 네이버 쇼핑검색광고 전문가입니다.

[1단계: 상품 파악]
대표 이미지와 상품명으로 이 상품의 정확한 카테고리를 파악하세요.
다른 카테고리 키워드는 절대 등록 금지. 예: 45박스 업소용냉장고 → 반찬냉장고 등록 금지.

[2단계: 키워드 선정 원칙 - 가장 중요]
★ 핵심 원칙: 실제 소비자가 네이버에서 검색하는 단어만 사용. 스펙·모델 정보는 절대 키워드로 쓰지 않음.

▶ 메인 키워드 (5~8개 필수):
- 카테고리명 그대로: "업소용냉장고", "영업용냉장고", "식당냉장고", "소형업소냉장고"
- 용도+카테고리: "식당용냉장고", "가게냉장고", "편의점냉장고"
- 검색량 높은 순으로 선정. 연관검색어 파일에서 월 검색량 반드시 확인.
- 절대 금지: "45박스", "LG", "올냉장", 숫자단위, 모델번호 단독 사용

▶ 서브 키워드 (4~6개 필수):
- 브랜드+카테고리 조합: "스타리온냉장고", "스타리온업소용냉장고"
- 좁은 범위 검색어: "45리터업소용냉장고", "소형업소용냉장고 추천"
- 반드시 실제 사람이 검색하는 완전한 단어 조합이어야 함
- 절대 금지: "45박스" 단독, "LG" 단독, "올냉장" 단독, 스펙 단독 숫자

[3단계: 수치]
※ 제외 키워드는 코드에서 자동 교차 계산됩니다. AI가 생성하지 않아도 됩니다.
- 노출 상품명 규칙 (최우선 - 반드시 지킬 것):
  ① 공백 포함 반드시 23~25자. 22자 이하 절대 금지.
  ② 26자 이상 절대 금지 (네이버 광고 노출 불가).
  ③ 단어 중간 자르기 절대 금지.
  ④ 글자 수가 부족하면 남은 공간을 계산해서 딱 맞는 키워드를 추가할 것.
     예: 20자 → 5자 남음 → "대형"(2자)+공백(1자)+"LG"(2자) = 5자 추가 → 25자 (O)
     예: 21자 → 4자 남음 → " LG대"는 단어 중간 자름(X) → " 대형"(3자) 추가 → 24자 (O)
  ⑤ 생성 후 반드시 글자수 검증: 공백 포함 한 글자씩 세기.
  예시:
    "스타리온 업소용냉장고 45박스 LG영업용" = 스(1)타(2)리(3)온(4) (5)업(6)소(7)용(8)냉(9)장(10)고(11) (12)4(13)5(14)박(15)스(16) (17)L(18)G(19)영(20)업(21)용(22) → 22자 (X, 짧음)
    "스타리온 업소용냉장고 45박스 LG영업용냉" → 단어 중간 자름 (X)
    "LG스타리온 업소용냉장고 45박스 영업용" = 24자 (O)
  productNames 배열에도 동일하게 23~25자로 서로 다른 조합 3개 생성.
- 총 하루 예산: 사용자 지정 금액이 있으면 그대로 사용. 없으면 평균 광고클릭비 × 25회 기준 현실적 산정.
- 입찰가 = 연관검색어 1위 PC 기준. 1번:×110%, 2번:×55%, 3번:×65%, 4번:×40%. 10원 단위.
- 하루예산 비율: 1번 47%, 2번 20%, 3번 20%, 4번 13%. 10원 단위.
- bidPc/bidMo: 연관검색어 파일의 광고 입찰가 1~3위 수치를 그대로 사용.

[분석 근거 주의]
- 파일에 실제 있는 수치만 사용 (월 검색량, 평균 광고클릭비 등)
- "X%로 구매의도가 명확합니다" 같은 근거 없는 문장 절대 금지
- reason 필드는 완전한 문장으로, 마침표로 끝낼 것

반드시 아래 JSON만 반환 (마크다운 없이, groups는 정확히 4개):
{
  "summary": {
    "mainKeywords": [{"kw": "업소용냉장고", "vol": 12000}, {"kw": "영업용냉장고", "vol": 8500}],
    "subKeywords": [{"kw": "스타리온냉장고", "vol": 3200}, {"kw": "스타리온업소용냉장고", "vol": 1800}],
    "bidTablePc": [1위입찰가, 2위입찰가, 3위입찰가],
    "bidTableMo": [1위입찰가, 2위입찰가, 3위입찰가],
    "totalBudget": 숫자,
    "reason": "전체 전략 요약."
  },
  "groups": [
    {
      "productName": "노출 상품명 23~25자",
      "productNames": ["상품명1 23~25자", "상품명2 23~25자"],
      "registerKeywords": "키워드1\n키워드2",
      "bidPc": [1위, 2위, 3위],
      "bidMo": [1위, 2위, 3위],
      "bidPrice": "숫자",
      "dailyBudget": "숫자"
    },
    { "productName": "2번 서브 23~25자", "productNames": ["..."], "registerKeywords": "...", "bidPc": [], "bidMo": [], "bidPrice": "숫자", "dailyBudget": "숫자" },
    { "productName": "3번 N+메인 23~25자", "productNames": ["..."], "registerKeywords": "...", "bidPc": [], "bidMo": [], "bidPrice": "숫자", "dailyBudget": "숫자" },
    { "productName": "4번 N+서브 23~25자", "productNames": ["..."], "registerKeywords": "...", "bidPc": [], "bidMo": [], "bidPrice": "숫자", "dailyBudget": "숫자" }
  ]
}`;

  const runAnalysis = async (extraComment = "") => {
    setLoading(true);
    setLoadingMsg("📂 파일 읽는 중...");
    try {
      const dataStr = await buildDataStr();
      setLoadingMsg("🤖 AI 분석 중...");
      const userPrompt = `상품명: ${productName}
브랜드: ${camp.brand || "스타리온"}
모델명: ${camp.model || ""}
메인키워드: ${camp.mainKw || ""}
하루 총 예산: ${camp.campTotalBudget ? camp.campTotalBudget + "원" : "지정 없음 (연관검색어 기준 자동산정)"}
${extraComment ? `\n추가 요청: ${extraComment}` : ""}

${dataStr}`;

      const raw = await callClaudeWithImage(userPrompt, systemPrompt, imagePreview);
      const jsonMatch = raw.match(/\{[\s\S]*\}/);
      if (!jsonMatch) throw new Error("JSON 파싱 실패");
      const result = JSON.parse(jsonMatch[0]);
      if (!result.groups || result.groups.length < 1) throw new Error("그룹 데이터 부족");
      // 정확히 4개로 맞추기
      while (result.groups.length < 4) result.groups.push({ ...result.groups[result.groups.length - 1] });
      result.groups = result.groups.slice(0, 4);
      setAnalysis({ ...result, rawData: dataStr });
      setStep("review");
      setComment("");
    } catch (err) {
      console.error(err);
      if (err.message === "API_KEY_MISSING") {
        showToast("API 키를 먼저 설정해주세요 (상단 🔑 버튼)");
      } else {
        showToast("분석 오류. 다시 시도해주세요.");
      }
    } finally {
      setLoading(false);
      setLoadingMsg("");
    }
  };

  const handleGenerate = () => {
    if (!productName.trim()) { showToast("상품명을 입력해주세요"); return; }
    if (!files.relKw && !files.adGroup) { showToast("연관 검색어 또는 광고 리포트 파일을 업로드해주세요"); return; }
    setStep("analyzing");
    runAnalysis();
  };

  const handleReanalyze = () => {
    if (!comment.trim()) { showToast("코멘트를 입력해주세요"); return; }
    runAnalysis(comment);
  };

  const handleApply = () => {
    if (!analysis) return;
    onApply(analysis.groups, analysis.summary);
    setStep("done");
    setOpen(false);
  };

  const handleReset = () => {
    setStep("input");
    setAnalysis(null);
    setComment("");
    setOpen(true);
  };

  const fileCount = Object.values(files).filter(Boolean).length;
  const canGenerate = productName.trim() && (files.relKw || files.adGroup);

  return (
    <div style={{ ...S.card, marginBottom: 14, border: step === "review" ? `1.5px solid ${NG}` : "1.5px solid #E0F2FE" }}>
      {/* Panel Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: open ? 16 : 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: step === "done" ? NG : "#0369A1" }}>
            {step === "done" ? "✅ 광고 세팅 완료" : "📥 광고 자료 입력 및 분석"}
          </span>
          {step === "review" && <span style={{ fontSize: 11, background: "#F0FDF4", color: NG, padding: "2px 8px", borderRadius: 10, fontWeight: 600 }}>검토 중</span>}
          {fileCount > 0 && step === "input" && !open && <span style={{ fontSize: 11, background: "#EFF6FF", color: "#2563EB", padding: "2px 8px", borderRadius: 10, fontWeight: 600 }}>파일 {fileCount}개</span>}
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {step === "done" && <button onClick={handleReset} style={{ fontSize: 11, color: "#444", background: "none", border: "0.5px solid #E5E7EB", borderRadius: 5, padding: "3px 8px", cursor: "pointer", fontWeight: 600 }}>↺ 다시 분석</button>}
          {step === "done" && !open && <button onClick={() => { setStep("input"); setOpen(true); }} style={{ fontSize: 11, color: "#0369A1", background: "#EFF6FF", border: "0.5px solid #BFDBFE", borderRadius: 5, padding: "3px 8px", cursor: "pointer", fontWeight: 600 }}>✏️ 편집</button>}
          {step === "review" && <button onClick={handleReset} style={{ fontSize: 11, color: "#333", background: "#F9FAFB", border: "0.5px solid #E5E7EB", borderRadius: 5, padding: "3px 8px", cursor: "pointer", fontWeight: 600 }}>← 되돌아가기</button>}
          <button onClick={() => setOpen(!open)} style={{ background: "none", border: "none", fontSize: 12, cursor: "pointer", color: "#444", fontWeight: 700 }}>
            {open ? "▲ 접기" : "▼ 펼치기"}
          </button>
        </div>
      </div>

      {open && (
        <>
          {/* STEP: input */}
          {(step === "input" || step === "analyzing" || step === "done" || step === "review") && step !== "review" && (
            <>
              <div style={{ marginBottom: 14 }}>
                <label style={S.label}>상품명 <span style={{ color: "#EF4444" }}>*</span></label>
                <input value={productName} onChange={e => setProductName(e.target.value)}
                  placeholder="예: 업소용 냉장고 냉동고 스타리온 45박스 LG A/S 영업용 대형 올냉장 올메탈 E45BAR" autoComplete="off"
                  style={{ ...S.input, fontSize: 13 }} />
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>
                {fileLabels.map(({ key, label, accept, icon, color, bg, border, isImage }) => (
                  <div key={key}>
                    <label style={{ ...S.label, color }}>{icon} {label}</label>
                    {files[key] ? (
                      <div style={{ border: `1.5px solid ${border}`, borderRadius: 8, padding: key === "image" ? 0 : "9px 12px", background: bg, overflow: "hidden", position: "relative" }}>
                        {isImage && imagePreview ? (
                          <>
                            <img src={imagePreview} alt="대표이미지" style={{ width: "100%", height: 80, objectFit: "cover", display: "block" }} />
                            <button onClick={() => removeFile(key)} style={{ position: "absolute", top: 4, right: 4, background: "rgba(0,0,0,0.5)", border: "none", color: "#fff", cursor: "pointer", fontSize: 12, borderRadius: 4, width: 20, height: 20, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
                          </>
                        ) : (
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 12px" }}>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontSize: 11, fontWeight: 600, color, wordBreak: "break-all" }}>{files[key].name}</div>
                              <div style={{ fontSize: 10, color: "#444", marginTop: 2 }}>{(files[key].size / 1024).toFixed(0)}KB</div>
                            </div>
                            <button onClick={() => removeFile(key)} style={{ background: "none", border: "none", color: "#444", cursor: "pointer", fontSize: 14, flexShrink: 0, marginLeft: 6 }}>✕</button>
                          </div>
                        )}
                      </div>
                    ) : (
                      <label style={{ border: `1.5px dashed ${border}`, borderRadius: 8, padding: "16px 12px", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", cursor: "pointer", background: bg, minHeight: 68 }}>
                        <span style={{ fontSize: 18, marginBottom: 3 }}>📁</span>
                        <span style={{ fontSize: 11, color: "#444", textAlign: "center" }}>클릭해서 업로드</span>
                        <span style={{ fontSize: 10, color: "#666", marginTop: 1 }}>xlsx</span>
                        <input type="file" accept={accept} onChange={e => handleFile(key, e)} style={{ display: "none" }} />
                      </label>
                    )}
                  </div>
                ))}
              </div>

              <button onClick={handleGenerate} disabled={!canGenerate || loading}
                style={{ width: "100%", padding: "13px 0", borderRadius: 10, border: "none", background: loading ? "#6EE7B7" : canGenerate ? NG : "#E5E7EB", color: canGenerate || loading ? "#fff" : "#9CA3AF", fontSize: 14, fontWeight: 800, cursor: canGenerate && !loading ? "pointer" : "not-allowed", letterSpacing: "0.04em" }}>
                {loading ? (
                  <span style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10 }}>
                    <span style={{ display: "inline-block", width: 15, height: 15, border: "2px solid #fff", borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
                    {loadingMsg}
                  </span>
                ) : "🔍 분석하기"}
              </button>
              {!canGenerate && !loading && (
                <div style={{ fontSize: 11, color: "#444", marginTop: 7, textAlign: "center" }}>상품명 입력 + 연관 검색어 또는 광고 리포트 파일 필요</div>
              )}
            </>
          )}

          {/* STEP: review */}
          {step === "review" && analysis && (
            <div>
              {/* Summary */}
              <div style={{ background: "#F0FDF4", border: "1.5px solid #D1FAE5", borderRadius: 10, padding: 16, marginBottom: 14 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#065F46", marginBottom: 12 }}>📊 AI 분석 결과</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 12 }}>
                  <div style={{ background: "#fff", borderRadius: 8, padding: "12px 14px" }}>
                    <div style={{ fontSize: 10, color: "#444", marginBottom: 8, fontWeight: 600 }}>메인 키워드 ({(analysis.summary?.mainKeywords || []).length}개)</div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                      {(analysis.summary?.mainKeywords || []).map((item, i) => {
                        const kw = typeof item === "object" ? item.kw : item;
                        const vol = typeof item === "object" ? item.vol : null;
                        return (
                          <div key={i} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "5px 10px", background: "#F0FDF4", borderRadius: 8, border: "1px solid #D1FAE5" }}>
                            <span style={{ fontSize: 12, fontWeight: 700, color: "#00873A" }}>{kw}</span>
                            {vol != null && <span style={{ fontSize: 11, color: "#333", fontWeight: 500 }}>월 {Number(vol).toLocaleString()}회</span>}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                  <div style={{ background: "#fff", borderRadius: 8, padding: "12px 14px" }}>
                    <div style={{ fontSize: 10, color: "#444", marginBottom: 8, fontWeight: 600 }}>서브 키워드 ({(analysis.summary?.subKeywords || []).length}개)</div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                      {(analysis.summary?.subKeywords || []).map((item, i) => {
                        const kw = typeof item === "object" ? item.kw : item;
                        const vol = typeof item === "object" ? item.vol : null;
                        return (
                          <div key={i} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "5px 10px", background: "#F5F3FF", borderRadius: 8, border: "1px solid #EDE9FE" }}>
                            <span style={{ fontSize: 12, fontWeight: 700, color: "#7C3AED" }}>{kw}</span>
                            {vol != null && <span style={{ fontSize: 11, color: "#333", fontWeight: 500 }}>월 {Number(vol).toLocaleString()}회</span>}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 12 }}>
                  <div style={{ background: "#fff", borderRadius: 8, padding: "8px 12px" }}>
                    <div style={{ fontSize: 10, color: "#444", marginBottom: 5, fontWeight: 600 }}>PC 입찰가 순위</div>
                    {(analysis.summary?.bidTablePc || [analysis.summary?.topBidPc]).filter(Boolean).map((v, i) => (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 2 }}>
                        <span style={{ color: "#444" }}>{i+1}위</span>
                        <span style={{ fontWeight: 700, color: "#111" }}>{fmt(v)}원</span>
                      </div>
                    ))}
                  </div>
                  <div style={{ background: "#fff", borderRadius: 8, padding: "8px 12px" }}>
                    <div style={{ fontSize: 10, color: "#444", marginBottom: 5, fontWeight: 600 }}>MO 입찰가 순위</div>
                    {(analysis.summary?.bidTableMo || [analysis.summary?.topBidMo]).filter(Boolean).map((v, i) => (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 2 }}>
                        <span style={{ color: "#444" }}>{i+1}위</span>
                        <span style={{ fontWeight: 700, color: "#111" }}>{fmt(v)}원</span>
                      </div>
                    ))}
                  </div>
                  <div style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", textAlign: "center" }}>
                    <div style={{ fontSize: 10, color: "#444", marginBottom: 5, fontWeight: 600 }}>총 예산</div>
                    <div style={{ fontSize: 14, fontWeight: 700, color: "#111" }}>{analysis.summary?.totalBudget ? `${fmt(analysis.summary.totalBudget)}원` : "—"}</div>
                    <div style={{ fontSize: 10, color: "#444", marginTop: 2 }}>일 기준</div>
                  </div>
                </div>
                {analysis.summary?.reason && (
                  <div style={{ background: "#fff", borderRadius: 8, padding: "12px 14px" }}>
                    <div style={{ fontSize: 10, color: "#444", fontWeight: 600, marginBottom: 8, letterSpacing: "0.04em" }}>💡 분석 근거</div>
                    {analysis.summary.reason.split(/[.。]/).filter(s => s.trim().length > 4).map((s, i) => (
                      <div key={i} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-start" }}>
                        <span style={{ fontSize: 11, fontWeight: 700, color: NG, flexShrink: 0, marginTop: 1 }}>·</span>
                        <span style={{ fontSize: 12, color: "#111", lineHeight: 1.65 }}>{s.trim()}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>



              {/* Comment & Action */}
              <div style={{ background: "#FFFBEB", border: "1.5px solid #FDE68A", borderRadius: 10, padding: 14, marginBottom: 12 }}>
                <label style={{ ...S.label, color: "#B45309" }}>💬 코멘트 (이상한 부분 있으면 작성 후 다시 분석)</label>
                <textarea value={comment} onChange={e => setComment(e.target.value)}
                  placeholder="예: 메인 키워드에 테이블냉장고 추가해줘 / 입찰가 더 낮게 / 제외 키워드에 V12BARD 추가해줘"
                  style={{ ...S.input, height: 70, resize: "none", marginBottom: 10, background: "#fff", borderColor: "#FDE68A" }} />
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={handleReanalyze} disabled={!comment.trim() || loading}
                    style={{ flex: 1, padding: "10px 0", borderRadius: 8, border: "1.5px solid #F59E0B", background: "#fff", fontSize: 13, fontWeight: 700, cursor: comment.trim() && !loading ? "pointer" : "not-allowed", color: comment.trim() ? "#B45309" : "#D1D5DB" }}>
                    {loading ? (
                      <span style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
                        <span style={{ display: "inline-block", width: 13, height: 13, border: "2px solid #B45309", borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
                        {loadingMsg}
                      </span>
                    ) : "↺ 다시 분석"}
                  </button>
                  <button onClick={handleApply} disabled={loading}
                    style={{ flex: 2, padding: "10px 0", borderRadius: 8, border: "none", background: loading ? "#9CA3AF" : NG, fontSize: 13, fontWeight: 800, cursor: !loading ? "pointer" : "not-allowed", color: "#fff", letterSpacing: "0.03em" }}>
                    ✅ 이상 없음 — 바로 적용하기
                  </button>
                </div>
              </div>
            </div>
          )}
        </>
      )}
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

// ── Four Group Grid ───────────────────────────────────────────────────────
function FourGroupGrid({ camp, onChange, showToast }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <CardPair camp={camp} idxA={0} idxB={1} onChange={onChange} showToast={showToast} />
      <CardPair camp={camp} idxA={2} idxB={3} onChange={onChange} showToast={showToast} />
    </div>
  );
}

// CardPair: 두 카드를 하나의 그리드로 묶어 섹션별 등높이 보장
function CardPair({ camp, idxA, idxB, onChange, showToast }) {
  const ga = camp.groups[idxA], gb = camp.groups[idxB];
  const metaA = GRP_META[idxA], metaB = GRP_META[idxB];

  // 각 카드의 공유 상태
  const [kwOpenA, setKwOpenA] = useState(true);
  const [kwOpenB, setKwOpenB] = useState(true);
  const [nameOpenA, setNameOpenA] = useState(true);
  const [nameOpenB, setNameOpenB] = useState(true);
  const [bidOpenA, setBidOpenA] = useState(true);
  const [bidOpenB, setBidOpenB] = useState(true);
  const [exclOpenA, setExclOpenA] = useState(true);
  const [exclOpenB, setExclOpenB] = useState(true);
  const [copiedA, setCopiedA] = useState(false);
  const [copiedB, setCopiedB] = useState(false);
  const [editBudgetA, setEditBudgetA] = useState(false);
  const [editBudgetB, setEditBudgetB] = useState(false);

  const updA = (field, val) => onChange(camp.id, ga.id, field, val);
  const updB = (field, val) => onChange(camp.id, gb.id, field, val);

  const copyExclA = () => { navigator.clipboard.writeText(ga.excludeKeywords); setCopiedA(true); setTimeout(() => setCopiedA(false), 2000); showToast("제외 키워드 복사됨!"); };
  const copyExclB = () => { navigator.clipboard.writeText(gb.excludeKeywords); setCopiedB(true); setTimeout(() => setCopiedB(false), 2000); showToast("제외 키워드 복사됨!"); };

  // 섹션 헤더 공통 컴포넌트
  const SHdr = ({ icon, title, sub, count, open, onToggle, right, color = "#111" }) => (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
      <button onClick={onToggle} style={{ background: "none", border: "none", cursor: "pointer", padding: 0, display: "flex", alignItems: "center", gap: 4 }}>
        <span style={{ fontSize: 12, fontWeight: 700, color }}>{icon} {title}</span>
        {sub && <span style={{ fontSize: 11, fontWeight: 500, color: "#555", marginLeft: 2 }}>{sub}</span>}
        <span style={{ fontSize: 11, color: "#555", marginLeft: 4 }}>{open ? "▲" : "▼"}</span>
      </button>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        {count != null && <span style={{ fontSize: 11, fontWeight: 600, color: "#111" }}>{count}개</span>}
        {right}
      </div>
    </div>
  );

  const renderBid = (g, meta, editBudget, setEditBudget, upd) => {
    if (g.bidPc?.length > 0 || g.bidMo?.length > 0) {
      const calcI = (i) => { const pc = n(g.bidPc?.[i]); const mo = n(g.bidMo?.[i]); if (pc > 0 && mo > 0) return Math.round((mo * 0.65 + pc * 0.35) / 10) * 10; return pc || mo; };
      const ranks = [0,1,2].filter(i => g.bidPc?.[i] || g.bidMo?.[i]);
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ background: "#fff", border: `1.5px solid ${meta.border}`, borderRadius: 8, padding: "10px 14px" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#111", marginBottom: 8 }}>통합 예상 입찰가</div>
            {ranks.map(i => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: i < ranks.length-1 ? 5 : 0 }}>
                <span style={{ fontSize: 12, color: "#111" }}>{i+1}위</span>
                <span style={{ fontSize: i===0 ? 16 : 13, fontWeight: 800, color: meta.color }}>{fmt(calcI(i))}원</span>
              </div>
            ))}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            <div style={{ background: "#fff", border: `1.5px solid ${meta.border}`, borderRadius: 8, padding: "9px 12px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#111", marginBottom: 6 }}>PC 입찰가</div>
              {(g.bidPc||[]).map((v,i) => <div key={i} style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:2 }}><span style={{color:"#111"}}>{i+1}위</span><span style={{fontWeight:700,color:meta.color}}>{fmt(v)}원</span></div>)}
            </div>
            <div style={{ background: "#fff", border: "1.5px solid #E5E7EB", borderRadius: 8, padding: "9px 12px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#111", marginBottom: 6 }}>MO 입찰가</div>
              {(g.bidMo||[]).map((v,i) => <div key={i} style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:2 }}><span style={{color:"#111"}}>{i+1}위</span><span style={{fontWeight:700,color:"#111"}}>{fmt(v)}원</span></div>)}
            </div>
          </div>
          <div style={{ background: "#fff", border: "1.5px solid #E5E7EB", borderRadius: 8, padding: "9px 12px", cursor: "pointer" }} onClick={() => setEditBudget(true)}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#111", marginBottom: 3 }}>추천 하루 예산</div>
            {editBudget ? (
              <div style={{ position: "relative" }}>
                <input autoFocus value={g.dailyBudget ? fmt(n(g.dailyBudget)) : ""} onChange={e => upd("dailyBudget", e.target.value.replace(/,/g,"").replace(/[^0-9]/g,""))} onBlur={() => setEditBudget(false)} inputMode="numeric" style={{ ...S.input, fontSize:15, fontWeight:700, padding:"2px 26px 2px 4px", border:"none", background:"transparent", width:"100%" }} />
                <span style={{ position:"absolute", right:2, top:"50%", transform:"translateY(-50%)", fontSize:11, color:"#111" }}>원</span>
              </div>
            ) : (
              <div style={{ display:"flex", alignItems:"baseline", gap:3 }}>
                <span style={{ fontSize:16, fontWeight:800, color:"#111" }}>{g.dailyBudget ? fmt(n(g.dailyBudget)) : "—"}</span>
                {g.dailyBudget && <span style={{ fontSize:11, color:"#111" }}>원</span>}
                <span style={{ fontSize:10, color:"#555", marginLeft:"auto" }}>✏️</span>
              </div>
            )}
          </div>
        </div>
      );
    }
    return (
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
        <div style={{ background:"#fff", border:`1.5px solid ${meta.border}`, borderRadius:8, padding:"9px 12px" }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#111", marginBottom:3 }}>기본 입찰가</div>
          <div style={{ fontSize:16, fontWeight:800, color:meta.color }}>{g.bidPrice ? fmt(n(g.bidPrice)) : "—"}{g.bidPrice && <span style={{fontSize:11}}>원</span>}</div>
        </div>
        <div style={{ background:"#fff", border:"1.5px solid #E5E7EB", borderRadius:8, padding:"9px 12px", cursor:"pointer" }} onClick={() => setEditBudget(true)}>
          <div style={{ fontSize:11, fontWeight:700, color:"#111", marginBottom:3 }}>하루 예산</div>
          {editBudget ? (
            <div style={{ position:"relative" }}>
              <input autoFocus value={g.dailyBudget ? fmt(n(g.dailyBudget)) : ""} onChange={e => upd("dailyBudget", e.target.value.replace(/,/g,"").replace(/[^0-9]/g,""))} onBlur={() => setEditBudget(false)} inputMode="numeric" style={{ ...S.input, fontSize:15, fontWeight:700, padding:"2px 26px 2px 4px", border:"none", background:"transparent", width:"100%" }} />
              <span style={{ position:"absolute", right:2, top:"50%", transform:"translateY(-50%)", fontSize:11 }}>원</span>
            </div>
          ) : (
            <div style={{ display:"flex", alignItems:"baseline", gap:3 }}>
              <span style={{ fontSize:16, fontWeight:800, color:"#111" }}>{g.dailyBudget ? fmt(n(g.dailyBudget)) : "—"}</span>
              {g.dailyBudget && <span style={{fontSize:11}}>원</span>}
              <span style={{ fontSize:10, color:"#555", marginLeft:"auto" }}>✏️</span>
            </div>
          )}
        </div>
      </div>
    );
  };

  const renderKwList = (g, meta) => {
    let kwItems = [];
    try { kwItems = JSON.parse(g.groupKeywords); } catch { kwItems = (g.groupKeywords||"").split(", ").filter(Boolean).map(k => ({ kw: k, vol: null })); }
    return (
      <div style={{ background: "#fff", border: `1.5px solid ${meta.border}`, borderRadius: 8, padding: "8px 12px", display: "flex", flexDirection: "column", gap: 5 }}>
        {kwItems.map((item,i) => (
          <div key={i} style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <span style={{ fontSize:13, fontWeight:700, color:"#111" }}>{item.kw}</span>
            {item.vol != null && <span style={{ fontSize:11, color:"#555", fontWeight:500 }}>월 {Number(item.vol).toLocaleString()}회</span>}
          </div>
        ))}
      </div>
    );
  };

  const renderProductNames = (g, meta, upd) => {
    const charCount = g.productName.length;
    const isOver = charCount > 25;
    return (
      <>
        <div style={{ position:"relative" }}>
          <input value={g.productName} onChange={e => upd("productName", e.target.value)} placeholder="23~25자 권장" maxLength={35}
            style={{ ...S.input, border:`1.5px solid ${isOver ? "#EF4444" : meta.border}`, paddingRight:52, fontSize:13, fontWeight:600, background:"#fff" }} />
          <span style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)", fontSize:11, fontWeight:700, color:isOver?"#EF4444":charCount>=22?"#F59E0B":"#555" }}>{charCount}/25</span>
        </div>
        {isOver && <div style={{ fontSize:11, color:"#EF4444", marginTop:4, fontWeight:700, background:"#FEF2F2", padding:"4px 8px", borderRadius:6 }}>⚠️ 25자 초과 — 네이버 노출 불가</div>}
        {(g.productNames||[]).slice(1).map((pn,i) => {
          const pc = pn.length;
          return (
            <div key={i} style={{ position:"relative", marginTop:6 }}>
              <input value={pn} onChange={e => { const arr=[...(g.productNames||[])]; arr[i+1]=e.target.value; upd("productNames",arr); }} maxLength={35}
                style={{ ...S.input, border:`1.5px solid ${pc>25?"#EF4444":meta.border}`, paddingRight:52, fontSize:13, fontWeight:600, opacity:0.85, background:"#fff" }} />
              <span style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)", fontSize:11, fontWeight:700, color:pc>25?"#EF4444":pc>=22?"#F59E0B":"#555" }}>{pc}/25</span>
            </div>
          );
        })}
      </>
    );
  };

  const renderSection = (sec, g, meta, open, onToggle, upd, editBudget, setEditBudget, copied, copyExcl, exclOpen) => {
    const rA = roasVal(g.revenue, g.cost);
    const badgeA = roasBadge(rA);
    switch(sec) {
      case "keyword":
        return (
          <div style={{ padding:"10px 14px", borderBottom:`1px solid ${meta.border}` }}>
            <SHdr icon={g.groupLabel==="메인"?"📌":"📎"} title={g.groupLabel==="메인"?"메인 키워드":"서브 키워드"} sub="(노출 예상 키워드)" count={kwCount(g.registerKeywords)} open={open} onToggle={onToggle} />
            {open && g.groupKeywords && <div style={{marginTop:8}}>{renderKwList(g,meta)}</div>}
          </div>
        );
      case "productName":
        return (
          <div style={{ padding:"10px 14px", borderBottom:`1px solid ${meta.border}` }}>
            <SHdr icon="🏷️" title="노출 상품명" open={open} onToggle={onToggle} />
            {open && <div style={{marginTop:8}}>{renderProductNames(g,meta,upd)}</div>}
          </div>
        );
      case "bid":
        return (
          <div style={{ padding:"10px 14px", borderBottom:`1px solid ${meta.border}` }}>
            <SHdr icon="💰" title="입찰가 & 예산" open={open} onToggle={onToggle} />
            {open && <div style={{marginTop:8}}>{renderBid(g,meta,editBudget,setEditBudget,upd)}</div>}
          </div>
        );
      case "excl":
        return (
          <div style={{ padding:"10px 14px" }}>
            <SHdr icon="🚫" title={`제외 키워드 (${kwCount(g.excludeKeywords)}개)`} open={exclOpen} onToggle={onToggle}
              right={<button onClick={copyExcl} style={{ background:copied?"#FEE2E2":"#FFF5F5", border:`1.5px solid ${copied?"#DC2626":"#FCA5A5"}`, borderRadius:6, padding:"4px 10px", fontSize:11, cursor:"pointer", fontWeight:700, color:copied?"#DC2626":"#EF4444" }}>{copied?"✓ 복사됨":"📋 복사"}</button>} />
            {exclOpen && <textarea value={g.excludeKeywords} onChange={e => upd("excludeKeywords",e.target.value)} placeholder="제외할 키워드를 줄바꿈으로 구분" style={{ ...S.input, marginTop:8, height:100, resize:"vertical", borderColor:"#FCA5A5", background:"#FFF0F0", fontSize:12, lineHeight:1.7, color:"#111" }} />}
            {n(g.clicks)>0 && (
              <div style={{ background:"#F9FAFB", borderRadius:8, padding:"8px 12px", display:"flex", justifyContent:"space-around", marginTop:8 }}>
                {[{label:"CPC",val:`${fmt(cpc(g.cost,g.clicks))}원`},{label:"전환율",val:`${(n(g.conversions)/n(g.clicks)*100).toFixed(1)}%`},{label:"ROAS",val:`${fmt(rA)}%`,hi:true}].map(m=>(
                  <div key={m.label} style={{textAlign:"center"}}>
                    <div style={{fontSize:10,color:"#555"}}>{m.label}</div>
                    <div style={{fontSize:13,fontWeight:700,color:m.hi?(rA>=1000?NG:"#EF4444"):"#111"}}>{m.val}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      default: return null;
    }
  };

  return (
    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
      {[{g:ga,meta:metaA,idx:idxA},{g:gb,meta:metaB,idx:idxB}].map(({g,meta,idx},ci) => {
        const isA = ci===0;
        const ew = isA?kwOpenA:kwOpenB, en=isA?nameOpenA:nameOpenB, eb=isA?bidOpenA:bidOpenB, ee=isA?exclOpenA:exclOpenB;
        const tw=isA?()=>setKwOpenA(v=>!v):()=>setKwOpenB(v=>!v), tn=isA?()=>setNameOpenA(v=>!v):()=>setNameOpenB(v=>!v), tb=isA?()=>setBidOpenA(v=>!v):()=>setBidOpenB(v=>!v), te=isA?()=>setExclOpenA(v=>!v):()=>setExclOpenB(v=>!v);
        const upd = isA?updA:updB, eb2=isA?editBudgetA:editBudgetB, seb=isA?setEditBudgetA:setEditBudgetB, cp=isA?copiedA:copiedB, cpe=isA?copyExclA:copyExclB;
        const r = roasVal(g.revenue, g.cost); const badge = roasBadge(r);
        return (
          <div key={idx} style={{ background:meta.bg, borderRadius:12, border:`1.5px solid ${meta.border}`, overflow:"hidden", boxShadow:"0 1px 6px rgba(0,0,0,0.06)", display:"flex", flexDirection:"column" }}>
            {/* 헤더 */}
            <div style={{ background:meta.color, padding:"10px 14px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <span style={{ fontSize:14, fontWeight:800, color:"#fff" }}>{meta.label}</span>
                <span style={{ fontSize:12, color:"rgba(255,255,255,0.9)", fontWeight:600 }}>{meta.sub}</span>
                {badge && n(g.clicks)>0 && <span style={{ fontSize:10, fontWeight:700, padding:"2px 7px", borderRadius:10, background:"rgba(255,255,255,0.2)", color:"#fff" }}>{badge.label}</span>}
              </div>
            </div>
            {renderSection("keyword",g,meta,ew,tw,upd,eb2,seb,cp,cpe,ee)}
            {renderSection("productName",g,meta,en,tn,upd,eb2,seb,cp,cpe,ee)}
            {renderSection("bid",g,meta,eb,tb,upd,eb2,seb,cp,cpe,ee)}
            {renderSection("excl",g,meta,ee,te,upd,eb2,seb,cp,cpe,ee)}
          </div>
        );
      })}
    </div>
  );
}

// ── API Key Modal ────────────────────────────────────────────────────────
function ApiKeyModal({ onSave, onClose }) {
  const [key, setKey] = useState(localStorage.getItem("forcool_api_key") || "");
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#ffffff", borderRadius: 14, width: 440, maxWidth: "92vw", boxShadow: "0 20px 60px rgba(0,0,0,0.3)", color: "#111" }}>
        <div style={{ background: "#111", padding: "14px 20px", display: "flex", justifyContent: "space-between", alignItems: "center", borderRadius: "12px 12px 0 0" }}>
          <span style={{ fontWeight: 700, fontSize: 14, color: "#fff" }}>🔑 Claude API 키 설정</span>
          <button onClick={onClose} style={{ background: "rgba(255,255,255,0.15)", border: "none", borderRadius: 5, color: "#fff", width: 26, height: 26, cursor: "pointer", fontSize: 14 }}>✕</button>
        </div>
        <div style={{ padding: 20, background: "#ffffff" }}>
          <div style={{ background: "#FFFBEB", border: "1px solid #FDE68A", borderRadius: 8, padding: "10px 14px", marginBottom: 16, fontSize: 12, color: "#92400E", lineHeight: 1.6 }}>
            ⚠️ API 키는 브라우저 로컬스토리지에만 저장됩니다.<br/>
            Anthropic Console → API Keys에서 발급받으세요.
          </div>
          <label style={{ fontSize: 11, fontWeight: 700, color: "#111", display: "block", marginBottom: 6 }}>API 키 (sk-ant-...)</label>
          <input
            value={key}
            onChange={e => setKey(e.target.value)}
            placeholder="sk-ant-api03-..."
            type="password"
            style={{ width: "100%", border: "1.5px solid #E5E7EB", borderRadius: 8, padding: "9px 12px", fontSize: 13, outline: "none", boxSizing: "border-box", background: "#FFFFFF", color: "#111827", fontFamily: "monospace", marginBottom: 16, WebkitTextFillColor: "#111827", WebkitBoxShadow: "0 0 0px 1000px #FFFFFF inset" }}
          />
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={onClose} style={{ flex: 1, padding: "10px 0", borderRadius: 8, border: "1px solid #E5E7EB", background: "#f9fafb", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#111" }}>취소</button>
            <button onClick={() => { if (key.trim()) onSave(key.trim()); }} disabled={!key.trim()}
              style={{ flex: 2, padding: "10px 0", borderRadius: 8, border: "none", background: key.trim() ? "#111" : "#E5E7EB", fontSize: 13, fontWeight: 700, cursor: key.trim() ? "pointer" : "not-allowed", color: key.trim() ? "#fff" : "#9CA3AF" }}>저장</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Camp Modal ────────────────────────────────────────────────────────────
function CampModal({ camp, onSave, onClose }) {
  const [num, setNum] = useState(camp?.number || "");
  const [adType, setAdType] = useState(camp?.adType || "쇼검");
  const [brand, setBrand] = useState(camp?.brand || "스타리온");
  const [model, setModel] = useState(camp?.model || "");
  const [mainKw, setMainKw] = useState(camp?.mainKw || "");
  const [campTotalBudget, setCampTotalBudget] = useState(camp?.campTotalBudget || "");
  const [brandCustom, setBrandCustom] = useState("");
  const [submitted, setSubmitted] = useState(false);

  const BRANDS = ["스타리온", "포쿨", "프레스코", "유니", "직접입력"];
  const AD_TYPES = ["쇼검", "파워"];
  const resolvedBrand = brand === "직접입력" ? brandCustom : brand;
  const preview = buildCampName(adType, resolvedBrand, model, mainKw);

  // 필수 필드 검증
  const errors = {
    num: !num.trim(),
    model: !model.trim(),
    mainKw: !mainKw.trim(),
    campTotalBudget: !campTotalBudget.trim(),
    brandCustom: brand === "직접입력" && !brandCustom.trim(),
  };
  const hasError = Object.values(errors).some(Boolean);

  const errStyle = (key) => submitted && errors[key]
    ? { border: "1.5px solid #EF4444", background: "#FFF5F5" }
    : {};

  const handleSave = () => {
    setSubmitted(true);
    if (hasError) return;
    onSave({ number: num.trim(), adType, brand: resolvedBrand, model: model.trim(), mainKw: mainKw.trim(), campTotalBudget: campTotalBudget.replace(/,/g, "").trim() });
  };

  // 예산 입력 시 쉼표 자동 포맷
  const handleBudgetChange = (e) => {
    const raw = e.target.value.replace(/,/g, "").replace(/[^0-9]/g, "");
    setCampTotalBudget(raw ? Number(raw).toLocaleString() : "");
  };

  const errMsg = (key, msg) => submitted && errors[key]
    ? <div style={{ fontSize: 11, color: "#EF4444", marginTop: 3, fontWeight: 600 }}>⚠️ {msg}</div>
    : null;

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#fff", borderRadius: 16, width: 480, maxWidth: "94vw", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ background: NG, padding: "16px 22px", display: "flex", justifyContent: "space-between", alignItems: "center", borderRadius: "14px 14px 0 0" }}>
          <span style={{ fontWeight: 700, fontSize: 15, color: "#fff" }}>{camp ? "캠페인 편집" : "새 캠페인 추가"}</span>
          <button onClick={onClose} style={{ background: "rgba(255,255,255,0.2)", border: "none", borderRadius: 6, color: "#fff", width: 28, height: 28, cursor: "pointer", fontSize: 15, fontWeight: 700 }}>✕</button>
        </div>
        <div style={{ padding: "22px 22px 0" }}>
          <div style={{ background: "#F0FDF4", border: "1.5px solid #D1FAE5", borderRadius: 10, padding: "12px 16px", marginBottom: 18 }}>
            <div style={{ fontSize: 10, color: "#444", marginBottom: 4, fontWeight: 600 }}>자동 생성 캠페인명</div>
            <div style={{ fontSize: 16, fontWeight: 700, color: "#111", wordBreak: "break-all" }}>{preview || "—"}</div>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={S.label}>캠페인 번호 <span style={{ color: "#EF4444" }}>*</span></label>
            <input value={num} onChange={e => setNum(e.target.value)} placeholder="예: 012" style={{ ...S.input, ...errStyle("num") }} />
            {errMsg("num", "캠페인 번호를 입력하세요")}
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={S.label}>광고 유형</label>
            <div style={{ display: "flex", gap: 8 }}>
              {AD_TYPES.map(t => <button key={t} onClick={() => setAdType(t)} style={{ ...S.btn(adType === t, "#2563EB"), flex: 1, padding: "9px 0", fontSize: 13 }}>{t}</button>)}
            </div>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={S.label}>브랜드</label>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {BRANDS.map(b => <button key={b} onClick={() => setBrand(b)} style={{ ...S.btn(brand === b, "#7C3AED"), padding: "7px 12px", fontSize: 12 }}>{b}</button>)}
            </div>
            {brand === "직접입력" && (
              <>
                <input value={brandCustom} onChange={e => setBrandCustom(e.target.value)} placeholder="브랜드명 직접 입력" style={{ ...S.input, marginTop: 8, ...errStyle("brandCustom") }} />
                {errMsg("brandCustom", "브랜드명을 입력하세요")}
              </>
            )}
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={S.label}>모델명 <span style={{ color: "#EF4444" }}>*</span></label>
            <input value={model} onChange={e => setModel(e.target.value)} placeholder="예: E45BAR" style={{ ...S.input, ...errStyle("model") }} />
            {errMsg("model", "모델명을 입력하세요")}
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={S.label}>메인 키워드 <span style={{ color: "#EF4444" }}>*</span></label>
            <input value={mainKw} onChange={e => setMainKw(e.target.value)} placeholder="예: 업소용냉장고" style={{ ...S.input, ...errStyle("mainKw") }} />
            {errMsg("mainKw", "메인 키워드를 입력하세요")}
          </div>
          <div style={{ marginBottom: 22 }}>
            <label style={S.label}>하루 총 광고 예산 <span style={{ color: "#EF4444" }}>*</span></label>
            <div style={{ position: "relative" }}>
              <input value={campTotalBudget} onChange={handleBudgetChange} inputMode="numeric" placeholder="예: 280,000" style={{ ...S.input, paddingRight: 28, ...errStyle("campTotalBudget") }} />
              <span style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", fontSize: 11, color: "#444" }}>원</span>
            </div>
            {errMsg("campTotalBudget", "하루 총 광고 예산을 입력하세요")}
          </div>
        </div>
        <div style={{ padding: "0 22px 22px", display: "flex", gap: 8 }}>
          <button onClick={onClose} style={{ flex: 1, padding: "11px 0", borderRadius: 8, border: "1.5px solid #E5E7EB", background: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", color: "#333" }}>취소</button>
          <button onClick={handleSave}
            style={{ flex: 2, padding: "11px 0", borderRadius: 8, border: "none", background: NG, fontSize: 13, fontWeight: 700, cursor: "pointer", color: "#fff" }}>
            {camp ? "저장" : "캠페인 추가"}
          </button>
        </div>
      </div>
    </div>
  );
}
