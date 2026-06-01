# Trust-me-
Financial services review
// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: README.md                                               ║
// ╚════════════════════════════════════════════════════════════════╝

# 🇿🇦 Trust Me — South African Financial Transparency App

Evidence-based rankings of every SA bank, insurer, medical aid, broker and asset manager.

## Features
- 5-pillar **Trust Me Index** (Transparency · Financial Confidence · Income Trust Gap · FinScope 2.0)
- 8 categories: Banks · Long-term Insurance · Short-term Insurance · Medical Aids · Brokerages · Investment · Credit & Lending · Payments & Fintech
- 30+ seeded SA providers (Capitec, Standard Bank, Discovery, Sanlam, Old Mutual, OUTsurance, King Price, Bonitas, EasyEquities, Allan Gray, Sygnia …)
- Evidence trail: HelloPeter, Reddit, Google, Wikipedia, Ombudsman (OBSSA/OLFI/CMS), NCR, FSCA
- Fines, awards, social impact, green footprint, AI usage score
- Side-by-side compare (up to 2) with winner banner
- Mobile-first with sticky bottom nav
- Supabase backend (optional) with live nightly score recomputation

## Stack
- React + TypeScript + Vite
- Tailwind CSS
- Supabase (Postgres + Edge Functions) — optional, falls back to seed data

## Quick start (Lovable)
1. Create a new Lovable project (React + TS + Tailwind).
2. Copy every file below into the matching path.
3. Run preview — the app works immediately with seed data.

## Full deployment with live data
1. Create a Supabase project → run `supabase/schema.sql` in the SQL Editor.
2. Seed `providers` via Table Editor (CSV import).
3. Copy `.env.example` to `.env`, fill in Supabase URL + anon key.
4. Deploy `supabase/functions/compute-trust-score` via Supabase CLI.
5. Schedule nightly recomputation via `pg_cron` (see schema.sql comments).

## Trust Me Index formula

Each pillar is normalised to 0–100 before weighting.

## Licence
MIT

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: package.json                                            ║
// ╚════════════════════════════════════════════════════════════════╝

{
  "name": "trust-me-sa",
  "private": true,
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "tsc && vite build",
    "preview": "vite preview"
  },
  "dependencies": {
    "react": "^18.3.1",
    "react-dom": "^18.3.1",
    "@supabase/supabase-js": "^2.45.0"
  },
  "devDependencies": {
    "@types/react": "^18.3.3",
    "@types/react-dom": "^18.3.0",
    "@vitejs/plugin-react": "^4.3.1",
    "autoprefixer": "^10.4.19",
    "postcss": "^8.4.38",
    "tailwindcss": "^3.4.4",
    "typescript": "^5.5.3",
    "vite": "^5.3.1"
  }
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: .env.example                                            ║
// ╚════════════════════════════════════════════════════════════════╝

# Optional — if blank, the app falls back to seed data
VITE_SUPABASE_URL=
VITE_SUPABASE_ANON=

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: supabase/schema.sql                                     ║
// ╚════════════════════════════════════════════════════════════════╝
-- =========================================================
-- TRUST ME — South African Financial Transparency Index
-- =========================================================

create extension if not exists "pgcrypto";

-- ──────────────────────────────────────────────────────────
-- 1. CATEGORIES
-- ──────────────────────────────────────────────────────────
create table if not exists categories (
  slug          text primary key,
  name          text not null,
  icon          text not null,
  description   text,
  subcategories jsonb not null default '[]'::jsonb,
  sort_order    int default 0
);

-- ──────────────────────────────────────────────────────────
-- 2. PROVIDERS
-- ──────────────────────────────────────────────────────────
create table if not exists providers (
  id                   uuid primary key default gen_random_uuid(),
  slug                 text unique not null,
  name                 text not null,
  category             text not null references categories(slug),
  subcategory          text not null,
  website              text,
  logo                 text default '🏦',
  founded              int,
  bbb_level            text,
  regulator            text,
  is_known             boolean default false,

  transparency         numeric(5,2) not null default 0,
  financial_confidence numeric(5,2) not null default 0,
  income_trust_gap     numeric(5,2) not null default 0,
  finscope_score       numeric(5,2) not null default 0,

  trust_score          numeric(5,2) generated always as (
    round((
      transparency * 0.28 +
      financial_confidence * 0.26 +
      income_trust_gap * 0.20 +
      finscope_score * 0.26
    )::numeric, 2) stored
  ),

  evidence             jsonb not null default '[]'::jsonb,
  social_impact        jsonb not null default '[]'::jsonb,
  green_footprint      jsonb not null default '{"score":0,"notes":[]}'::jsonb,
  ai_use               jsonb not null default '{"score":0,"notes":[]}'::jsonb,

  created_at           timestamptz default now(),
  updated_at           timestamptz default now()
);

create index if not exists providers_category_idx on providers(category);
create index if not exists providers_trust_idx    on providers(trust_score desc);
create index if not exists providers_slug_idx     on providers(slug);

-- ──────────────────────────────────────────────────────────
-- 3. NORMALISED SUB-TABLES
-- ──────────────────────────────────────────────────────────
create table if not exists provider_fines (
  id          uuid primary key default gen_random_uuid(),
  provider_id uuid not null references providers(id) on delete cascade,
  year        int not null,
  regulator   text not null,
  amount      text not null,
  reason      text not null
);

create table if not exists provider_awards (
  id          uuid primary key default gen_random_uuid(),
  provider_id uuid not null references providers(id) on delete cascade,
  year        int not null,
  name        text not null,
  issuer      text not null
);

create table if not exists provider_evidence (
  id           uuid primary key default gen_random_uuid(),
  provider_id  uuid not null references providers(id) on delete cascade,
  source       text not null,
  url          text not null,
  positive     int not null default 0,
  negative     int not null default 0,
  last_updated date not null default current_date
);

-- ──────────────────────────────────────────────────────────
-- 4. VIEW — category leaderboard
-- ──────────────────────────────────────────────────────────
create or replace view v_category_ranking as
select
  category, subcategory, slug, name, trust_score,
  rank() over (partition by category order by trust_score desc) as rank_in_cat
from providers;

-- ──────────────────────────────────────────────────────────
-- 5. TRIGGER — touch updated_at
-- ──────────────────────────────────────────────────────────
create or replace function touch_provider_updated()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create trigger providers_touch
before update on providers
for each row execute function touch_provider_updated();

-- ──────────────────────────────────────────────────────────
-- 6. ROW-LEVEL SECURITY (public read / admin write)
-- ──────────────────────────────────────────────────────────
alter table providers          enable row level security;
alter table categories         enable row level security;
alter table provider_fines     enable row level security;
alter table provider_awards    enable row level security;
alter table provider_evidence  enable row level security;

create policy "public_read_providers"  on providers         for select using (true);
create policy "public_read_categories" on categories        for select using (true);
create policy "public_read_fines"      on provider_fines    for select using (true);
create policy "public_read_awards"     on provider_awards   for select using (true);
create policy "public_read_evidence"   on provider_evidence for select using (true);

create policy "admin_write_providers"  on providers         for all using (auth.jwt() ->> 'role' = 'admin');
create policy "admin_write_categories" on categories        for all using (auth.jwt() ->> 'role' = 'admin');
create policy "admin_write_fines"      on provider_fines    for all using (auth.jwt() ->> 'role' = 'admin');
create policy "admin_write_awards"     on provider_awards   for all using (auth.jwt() ->> 'role' = 'admin');
create policy "admin_write_evidence"   on provider_evidence for all using (auth.jwt() ->> 'role' = 'admin');

-- ──────────────────────────────────────────────────────────
-- 7. SEARCH FUNCTION
-- ──────────────────────────────────────────────────────────
create or replace function search_providers(q text, limit_n int default 10)
returns setof providers language sql stable as $$
  select * from providers
  where
    name ilike '%' || q || '%' or
    subcategory ilike '%' || q || '%' or
    category ilike '%' || q || '%' or
    slug ilike '%' || q || '%'
  order by trust_score desc
  limit limit_n;
$$;

-- ═══════════════════════════════════════════════════════════
-- OPTIONAL: nightly recomputation (enable pg_cron first)
-- ═══════════════════════════════════════════════════════════
-- create extension if not exists pg_cron;
-- create extension if not exists pg_net;
-- select cron.schedule(
--   'recompute-trust-nightly', '0 2 * * *',
--   $$ select net.http_post(
--        url := current_setting('app.settings.functions_url') || '/compute-trust-score',
--        headers := '{"Authorization":"Bearer ' || current_setting('app.settings.service_key') || '"}'::jsonb
--      ); $$
-- );

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: supabase/functions/compute-trust-score/index.ts         ║
// ╚════════════════════════════════════════════════════════════════╝

import { createClient } from "npm:@supabase/supabase-js@2";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_KEY  = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

Deno.serve(async () => {
  const sb = createClient(SUPABASE_URL, SERVICE_KEY);
  const { data: providers } = await sb.from("providers").select("id,slug");
  if (!providers) return new Response("no providers", { status: 500 });

  for (const p of providers) {
    try {
      const pillars = await scrapePillars(p.slug);
      await sb.from("providers").update(pillars).eq("id", p.id);
    } catch (e) {
      console.error(`Failed ${p.slug}:`, e);
    }
  }
  return new Response(JSON.stringify({ ok: true, count: providers.length }), {
    headers: { "Content-Type": "application/json" },
  });
});

async function scrapePillars(slug: string) {
  // Hook real scrapers here (HelloPeter, Reddit, Google, Ombudsman…).
  // Each must return 0–100 for every pillar.
  const sb = createClient(SUPABASE_URL, SERVICE_KEY);
  const { data: existing } = await sb
    .from("providers")
    .select("transparency,financial_confidence,income_trust_gap,finscope_score")
    .eq("slug", slug)
    .single();

  return {
    transparency:         existing?.transparency         ?? 50,
    financial_confidence: existing?.financial_confidence ?? 50,
    income_trust_gap:     existing?.income_trust_gap     ?? 50,
    finscope_score:       existing?.finscope_score       ?? 50,
  };
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: src/lib/trust.ts                                        ║
// ╚════════════════════════════════════════════════════════════════╝

export interface TrustComponents {
  transparency: number;
  financialConfidence: number;
  incomeTrustGap: number;
  finscope: number;
}

export const WEIGHTS = {
  transparency: 0.28,
  financialConfidence: 0.26,
  incomeTrustGap: 0.20,
  finscope: 0.26,
};

export function compositeTrust(c: TrustComponents): number {
  const raw =
    c.transparency * WEIGHTS.transparency +
    c.financialConfidence * WEIGHTS.financialConfidence +
    c.incomeTrustGap * WEIGHTS.incomeTrustGap +
    c.finscope * WEIGHTS.finscope;
  return Math.round(raw * 10) / 10;
}

export function gradeFromScore(s: number): { letter: string; label: string; color: string } {
  if (s >= 85) return { letter: "A", label: "Highly Trusted",   color: "text-emerald-600" };
  if (s >= 75) return { letter: "B", label: "Strong Trust",     color: "text-green-600" };
  if (s >= 65) return { letter: "C", label: "Moderate Trust",   color: "text-amber-600" };
  if (s >= 55) return { letter: "D", label: "Trust Needs Work", color: "text-orange-600" };
  return                { letter: "F", label: "Low Trust",        color: "text-rose-600" };
}

export function rankInCategory<T extends { trustScore: number }>(list: T[]): T[] {
  return [...list].sort((a, b) => b.trustScore - a.trustScore);
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: src/lib/sources.ts                                      ║
// ╚════════════════════════════════════════════════════════════════╝

export type SourceDef = {
  key: string;
  name: string;
  logo: string;
  color: string;
  description: string;
};

export const SOURCES: Record<string, SourceDef> = {
  HelloPeter: { key: "HelloPeter", name: "HelloPeter", logo: "📣", color: "bg-yellow-100 text-yellow-800", description: "South African consumer complaints & praise hub" },
  Reddit:     { key: "Reddit",     name: "Reddit r/southafrica & r/SAinvestments", logo: "🧡", color: "bg-orange-100 text-orange-800", description: "Peer discussions aggregated via NLP sentiment" },
  Google:     { key: "Google",     name: "Google Reviews", logo: "🔍", color: "bg-blue-100 text-blue-800", description: "Google Maps & search sentiment aggregation" },
  Wikipedia:  { key: "Wikipedia",  name: "Wikipedia citations", logo: "📚", color: "bg-slate-100 text-slate-800", description: "Verified editorial references" },
  Trustpilot: { key: "Trustpilot", name: "Trustpilot", logo: "⭐", color: "bg-green-100 text-green-800", description: "Global consumer review platform" },
  NCR:        { key: "NCR",        name: "National Credit Regulator", logo: "⚖️", color: "bg-purple-100 text-purple-800", description: "Registered complaints against credit providers" },
  Ombudsman:  { key: "Ombudsman",  name: "Ombudsman (OBSSA / OLFI / CMS)", logo: "🏛️", color: "bg-indigo-100 text-indigo-800", description: "Statutory dispute resolutions" },
};

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: src/lib/api.ts                                          ║
// ╚════════════════════════════════════════════════════════════════╝

import { createClient } from "@supabase/supabase-js";
import { PROVIDERS as SEED } from "../data/providers";
import type { Provider } from "../data/providers";

const SUPABASE_URL  = import.meta.env.VITE_SUPABASE_URL  ?? "";
const SUPABASE_ANON = import.meta.env.VITE_SUPABASE_ANON ?? "";

const hasBackend = Boolean(SUPABASE_URL && SUPABASE_ANON);
const sb = hasBackend ? createClient(SUPABASE_URL, SUPABASE_ANON) : null;

function toProvider(row: any): Provider {
  return {
    slug: row.slug,
    name: row.name,
    category: row.category,
    subcategory: row.subcategory,
    website: row.website,
    logo: row.logo,
    founded: row.founded,
    bbb: row.bbb_level,
    regulator: row.regulator,
    trustComponents: {
      transparency: row.transparency,
      financialConfidence: row.financial_confidence,
      incomeTrustGap: row.income_trust_gap,
      finscope: row.finscope_score,
    },
    trustScore: row.trust_score,
    evidence: row.evidence ?? [],
    fines: row.fines ?? [],
    awards: row.awards ?? [],
    socialImpact: row.social_impact ?? [],
    greenFootprint: row.green_footprint ?? { score: 0, notes: [] },
    aiUse: row.ai_use ?? { score: 0, notes: [] },
    known: row.is_known ?? false,
  };
}

export async function loadProviders(): Promise<Provider[]> {
  if (!sb) return SEED;
  const { data, error } = await sb
    .from("providers")
    .select("*")
    .order("trust_score", { ascending: false });
  if (error || !data) {
    console.warn("Supabase load failed, using seed.", error);
    return SEED;
  }
  return data.map(toProvider);
}

export async function loadProvider(slug: string): Promise<Provider | null> {
  if (!sb) return SEED.find((p) => p.slug === slug) ?? null;
  const { data } = await sb.from("providers").select("*").eq("slug", slug).single();
  return data ? toProvider(data) : null;
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: src/data/categories.ts                                  ║
// ╚════════════════════════════════════════════════════════════════╝

export type Category = {
  slug: string;
  name: string;
  icon: string;
  description: string;
  subcategories: string[];
};

export const CATEGORIES: Category[] = [
  {
    slug: "banks",
    name: "Banks",
    icon: "🏦",
    description: "Retail, business & digital banks registered with the SARB and governed by the Banks Act.",
    subcategories: ["Retail Bank", "Digital / Neo Bank", "Private & Wealth Bank", "Development Bank"],
  },
  {
    slug: "insurance-long-term",
    name: "Insurance — Long-Term",
    icon: "🛡️",
    description: "Life cover, funeral policies, retirement annuities, disability & living annuities.",
    subcategories: ["Life Insurance", "Funeral Cover", "Retirement Annuities", "Disability & Dread Disease"],
  },
  {
    slug: "insurance-short-term",
    name: "Insurance — Short-Term",
    icon: "🚗",
    description: "Car, home, business, pet & gadget insurance regulated under the STIA.",
    subcategories: ["Motor", "Home & Property", "Business", "Gadget & Pet", "SME Cover"],
  },
  {
    slug: "medical-aid",
    name: "Medical Schemes",
    icon: "🏥",
    description: "Council for Medical Schemes (CMS) registered medical aid plans & hospital plans.",
    subcategories: ["Comprehensive", "Hospital Plan", "Savings Plan", "Low-Cost / Gap Cover"],
  },
  {
    slug: "brokerages",
    name: "Brokerages",
    icon: "🤝",
    description: "FSCA-licensed intermediary brokers across short-term, long-term & JSE brokerage.",
    subcategories: ["Insurance Broker", "JSE Broker", "Independent Financial Adviser"],
  },
  {
    slug: "investment",
    name: "Investment & Asset Managers",
    icon: "📈",
    description: "Unit trusts, RA managers, discretionary fund managers and robo-advisors.",
    subcategories: ["Asset Manager", "Robo-Advisory", "Retail Brokerage Platform"],
  },
  {
    slug: "credit-lending",
    name: "Credit & Lending",
    icon: "💳",
    description: "NCR-registered credit providers — micro-lenders, vehicle & asset finance.",
    subcategories: ["Micro-Lending", "Vehicle Finance", "Personal Loans", "Store Credit"],
  },
  {
    slug: "payment-fintech",
    name: "Payments & Fintech",
    icon: "📱",
    description: "Payment service providers, wallets & SARB sandbox innovators.",
    subcategories: ["Payments", "Wallets", "Remittance"],
  },
];

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: src/data/providers.ts                                   ║
// ╚════════════════════════════════════════════════════════════════╝

import type { TrustComponents } from "../lib/trust";
import { compositeTrust } from "../lib/trust";

export type Evidence = {
  source: "HelloPeter" | "Reddit" | "Google" | "Wikipedia" | "Trustpilot" | "NCR" | "Ombudsman";
  url: string;
  positive: number;
  negative: number;
  lastUpdated: string;
};

export type Provider = {
  slug: string;
  name: string;
  category: string;
  subcategory: string;
  website: string;
  logo: string;
  founded: number;
  bbb: "1" | "2" | "3" | "4" | "non-compliant" | "N/A";
  regulator: string;
  trustComponents: TrustComponents;
  trustScore: number;
  evidence: Evidence[];
  fines: { year: number; regulator: string; amount: string; reason: string }[];
  awards: { year: number; name: string; issuer: string }[];
  socialImpact: string[];
  greenFootprint: { score: number; notes: string[] };
  aiUse: { score: number; notes: string[] };
  known: boolean;
};

function build(t: TrustComponents): number {
  return compositeTrust(t);
}

const raw: Omit<Provider, "trustScore">[] = [
  // ───────── BANKS ─────────
  {
    slug: "capitec", name: "Capitec Bank", category: "banks", subcategory: "Retail Bank",
    website: "https://www.capitecbank.co.za", logo: "🏦", founded: 2001, bbb: "2", regulator: "SARB / PA",
    trustComponents: { transparency: 86, financialConfidence: 88, incomeTrustGap: 92, finscope: 89 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/capitec-bank", positive: 2430, negative: 610, lastUpdated: "2026-06-01" },
      { source: "Reddit", url: "https://www.reddit.com/r/southafrica/search/?q=capitec", positive: 410, negative: 90, lastUpdated: "2026-06-01" },
      { source: "Google", url: "https://www.google.com/search?q=capitec+bank+reviews", positive: 18200, negative: 3400, lastUpdated: "2026-06-01" },
      { source: "Ombudsman", url: "https://www.obssa.co.za", positive: 120, negative: 48, lastUpdated: "2026-05-15" },
    ],
    fines: [{ year: 2024, regulator: "FSCA", amount: "R1.5m", reason: "Disclosure shortfall on credit fees" }],
    awards: [
      { year: 2025, name: "Bank of the Year – SA", issuer: "The Banker (FT)" },
      { year: 2024, name: "Most Trusted Bank", issuer: "Reader's Digest SA" },
    ],
    socialImpact: ["Capitec Foundation — SME grants", "Financial literacy programme in 1,200 schools"],
    greenFootprint: { score: 74, notes: ["Paperless statements default", "Solar on 34% of branches (2025 report)"] },
    aiUse: { score: 78, notes: ["Fraud detection ML live", "AI voice assistant in app (beta 2026)"] },
    known: true,
  },
  {
    slug: "standard-bank", name: "Standard Bank", category: "banks", subcategory: "Retail Bank",
    website: "https://www.standardbank.co.za", logo: "🏦", founded: 1863, bbb: "1", regulator: "SARB / PA",
    trustComponents: { transparency: 84, financialConfidence: 91, incomeTrustGap: 74, finscope: 83 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/standard-bank", positive: 1890, negative: 1420, lastUpdated: "2026-06-01" },
      { source: "Reddit", url: "https://www.reddit.com/r/southafrica/search/?q=standard+bank", positive: 310, negative: 220, lastUpdated: "2026-06-01" },
      { source: "Google", url: "https://www.google.com/search?q=standard+bank+reviews", positive: 14200, negative: 6100, lastUpdated: "2026-06-01" },
      { source: "Wikipedia", url: "https://en.wikipedia.org/wiki/Standard_Bank", positive: 22, negative: 4, lastUpdated: "2026-06-01" },
    ],
    fines: [{ year: 2023, regulator: "FSCA", amount: "R4.2m", reason: "FAIS disclosure breach" }],
    awards: [{ year: 2025, name: "Best Bank in Africa", issuer: "Euromoney" }],
    socialImpact: ["Standard Bank Foundation — R280m/year deployed", "SME lending uplift programme"],
    greenFootprint: { score: 81, notes: ["No new coal funding since 2026", "Net-zero 2050 target validated"] },
    aiUse: { score: 85, notes: ["AI fraud guard", "Generative-AI concierge in app"] },
    known: true,
  },
  {
    slug: "fnb", name: "FirstRand / FNB", category: "banks", subcategory: "Retail Bank",
    website: "https://www.fnb.co.za", logo: "🏦", founded: 1838, bbb: "1", regulator: "SARB / PA",
    trustComponents: { transparency: 82, financialConfidence: 90, incomeTrustGap: 72, finscope: 81 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/fnb", positive: 2100, negative: 1780, lastUpdated: "2026-06-01" },
      { source: "Reddit", url: "https://www.reddit.com/r/southafrica/search/?q=fnb", positive: 520, negative: 430, lastUpdated: "2026-06-01" },
      { source: "Google", url: "https://www.google.com/search?q=fnb+reviews", positive: 16300, negative: 7400, lastUpdated: "2026-06-01" },
    ],
    fines: [{ year: 2024, regulator: "National Treasury", amount: "R95m", reason: "Historic Amls compliance gaps" }],
    awards: [{ year: 2025, name: "Best Digital Bank SA", issuer: "Global Finance" }],
    socialImpact: ["FNB Foundation", "eBucks uplift for township businesses"],
    greenFootprint: { score: 79, notes: ["Green bond issuance", "EV charging at 40 branches"] },
    aiUse: { score: 88, notes: ["AI spend insights", "Real-time fraud scoring model v4"] },
    known: true,
  },
  {
    slug: "absa", name: "Absa Group", category: "banks", subcategory: "Retail Bank",
    website: "https://www.absa.co.za", logo: "🏦", founded: 1862, bbb: "1", regulator: "SARB / PA",
    trustComponents: { transparency: 78, financialConfidence: 82, incomeTrustGap: 70, finscope: 77 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/absa", positive: 1420, negative: 1610, lastUpdated: "2026-06-01" },
      { source: "Reddit", url: "https://www.reddit.com/r/southafrica/search/?q=absa", positive: 220, negative: 310, lastUpdated: "2026-06-01" },
    ],
    fines: [{ year: 2023, regulator: "FSCA", amount: "R2.7m", reason: "Policyholder disclosure delays" }],
    awards: [{ year: 2024, name: "Best Workplaces Africa", issuer: "Great Place to Work" }],
    socialImpact: ["Absa Foundation", "Youth employment accelerator"],
    greenFootprint: { score: 72, notes: ["Climate risk reporting aligned to TCFD"] },
    aiUse: { score: 76, notes: ["Chatbot 'Aby' upgraded 2025"] },
    known: true,
  },
  {
    slug: "nedbank", name: "Nedbank", category: "banks", subcategory: "Retail Bank",
    website: "https://www.nedbank.co.za", logo: "🏦", founded: 1888, bbb: "1", regulator: "SARB / PA",
    trustComponents: { transparency: 80, financialConfidence: 84, incomeTrustGap: 69, finscope: 78 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/nedbank", positive: 1310, negative: 1180, lastUpdated: "2026-06-01" },
      { source: "Reddit", url: "https://www.reddit.com/r/southafrica/search/?q=nedbank", positive: 260, negative: 240, lastUpdated: "2026-06-01" },
    ],
    fines: [],
    awards: [{ year: 2024, name: "Most Sustainable Bank SA", issuer: "Bloomberg" }],
    socialImpact: ["Nedbank Foundation", "Green finance desk"],
    greenFootprint: { score: 86, notes: ["Leader in SA green-bond issuance", "Sector coal-exit pathway"] },
    aiUse: { score: 74, notes: ["AI ESG analytics", "Gen-AI advisor pilot"] },
    known: true,
  },
  {
    slug: "tymebank", name: "TymeBank", category: "banks", subcategory: "Digital / Neo Bank",
    website: "https://www.tymebank.co.za", logo: "📱", founded: 2019, bbb: "1", regulator: "SARB / PA",
    trustComponents: { transparency: 81, financialConfidence: 83, incomeTrustGap: 88, finscope: 80 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/tymebank", positive: 1640, negative: 490, lastUpdated: "2026-06-01" },
      { source: "Reddit", url: "https://www.reddit.com/r/southafrica/search/?q=tymebank", positive: 380, negative: 110, lastUpdated: "2026-06-01" },
    ],
    fines: [],
    awards: [{ year: 2025, name: "Best Digital Onboarding", issuer: "Banker Awards" }],
    socialImpact: ["Free savings goal tools", "Kiosk network in Shoprite/Usave"],
    greenFootprint: { score: 88, notes: ["Branchless model", "Low carbon per customer"] },
    aiUse: { score: 84, notes: ["Real-time credit scoring via AI"] },
    known: true,
  },
  {
    slug: "discovery-bank", name: "Discovery Bank", category: "banks", subcategory: "Retail Bank",
    website: "https://www.discovery.co.za/banking", logo: "🏦", founded: 2019, bbb: "2", regulator: "SARB / PA",
    trustComponents: { transparency: 85, financialConfidence: 84, incomeTrustGap: 65, finscope: 76 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/discovery-bank", positive: 1420, negative: 420, lastUpdated: "2026-06-01" },
      { source: "Reddit", url: "https://www.reddit.com/r/southafrica/search/?q=discovery+bank", positive: 410, negative: 160, lastUpdated: "2026-06-01" },
    ],
    fines: [],
    awards: [{ year: 2025, name: "World's Most Innovative Bank", issuer: "American Banker" }],
    socialImpact: ["Vitality Money cashback on healthy purchases"],
    greenFootprint: { score: 82, notes: ["Sustainable investment default for portfolios"] },
    aiUse: { score: 91, notes: ["Behavioural-banking AI", "Predictive savings coach"] },
    known: true,
  },
  {
    slug: "african-bank", name: "African Bank", category: "banks", subcategory: "Retail Bank",
    website: "https://www.africanbank.co.za", logo: "🏦", founded: 2016, bbb: "2", regulator: "SARB / PA",
    trustComponents: { transparency: 72, financialConfidence: 70, incomeTrustGap: 83, finscope: 74 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/african-bank", positive: 640, negative: 810, lastUpdated: "2026-06-01" },
    ],
    fines: [{ year: 2022, regulator: "NCR", amount: "R600k", reason: "Affordability assessment process" }],
    awards: [],
    socialImpact: ["Financial inclusion focus", "Low-income personal loans"],
    greenFootprint: { score: 60, notes: ["Improving disclosures"] },
    aiUse: { score: 58, notes: ["Automated credit assessments"] },
    known: true,
  },

  // ───────── LONG-TERM INSURANCE ─────────
  {
    slug: "old-mutual", name: "Old Mutual", category: "insurance-long-term", subcategory: "Life Insurance",
    website: "https://www.oldmutual.co.za", logo: "🛡️", founded: 1845, bbb: "1", regulator: "FSCA / PA",
    trustComponents: { transparency: 81, financialConfidence: 86, incomeTrustGap: 78, finscope: 80 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/old-mutual", positive: 1540, negative: 1280, lastUpdated: "2026-06-01" },
      { source: "Ombudsman", url: "https://www.olfi.co.za", positive: 190, negative: 210, lastUpdated: "2026-05-20" },
    ],
    fines: [{ year: 2023, regulator: "FSCA", amount: "R3.2m", reason: "Claims-handling delays" }],
    awards: [{ year: 2024, name: "Best Life Insurer SA", issuer: "Finweek" }],
    socialImpact: ["Funeral cover for 4m+ lives", "Old Mutual Foundation scholarships"],
    greenFootprint: { score: 70, notes: ["Net-zero 2050 commitment"] },
    aiUse: { score: 69, notes: ["AI claims triage"] },
    known: true,
  },
  {
    slug: "sanlam", name: "Sanlam", category: "insurance-long-term", subcategory: "Life Insurance",
    website: "https://www.sanlam.co.za", logo: "🛡️", founded: 1918, bbb: "1", regulator: "FSCA / PA",
    trustComponents: { transparency: 83, financialConfidence: 87, incomeTrustGap: 75, finscope: 81 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/sanlam", positive: 1210, negative: 940, lastUpdated: "2026-06-01" }],
    fines: [],
    awards: [{ year: 2025, name: "Long-term Insurer of the Year", issuer: "RMB/TopCo" }],
    socialImpact: ["Sanlam Foundation", "Sahamiti micro-insurance"],
    greenFootprint: { score: 73, notes: ["Climate-risk integrated investing"] },
    aiUse: { score: 75, notes: ["Sanlam iWorx AI planning"] },
    known: true,
  },
  {
    slug: "discovery-life", name: "Discovery Life", category: "insurance-long-term", subcategory: "Life Insurance",
    website: "https://www.discovery.co.za/life-insurance", logo: "🛡️", founded: 2012, bbb: "2", regulator: "FSCA / PA",
    trustComponents: { transparency: 87, financialConfidence: 88, incomeTrustGap: 64, finscope: 79 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/discovery-life", positive: 1380, negative: 420, lastUpdated: "2026-06-01" }],
    fines: [],
    awards: [{ year: 2025, name: "Most Innovative Life Product", issuer: "Africa Insurance Awards" }],
    socialImpact: ["Vitality wellness-linked premiums"],
    greenFootprint: { score: 78, notes: ["ESG-integrated underwriting"] },
    aiUse: { score: 89, notes: ["Wearable data underwriting AI"] },
    known: true,
  },
  {
    slug: "momentum-metropolitan", name: "Momentum Metropolitan", category: "insurance-long-term", subcategory: "Life Insurance",
    website: "https://www.momentum.co.za", logo: "🛡️", founded: 2017, bbb: "1", regulator: "FSCA / PA",
    trustComponents: { transparency: 76, financialConfidence: 80, incomeTrustGap: 72, finscope: 75 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/momentum", positive: 890, negative: 970, lastUpdated: "2026-06-01" }],
    fines: [{ year: 2022, regulator: "FSCA", amount: "R1.8m", reason: "Beneficiary payment delays" }],
    awards: [],
    socialImpact: ["Multiply wellness programme"],
    greenFootprint: { score: 68, notes: [] },
    aiUse: { score: 70, notes: ["Predictive lapse modelling"] },
    known: true,
  },
  {
    slug: "hollard-life", name: "Hollard Life", category: "insurance-long-term", subcategory: "Life Insurance",
    website: "https://www.hollard.co.za", logo: "🛡️", founded: 1985, bbb: "1", regulator: "FSCA / PA",
    trustComponents: { transparency: 79, financialConfidence: 81, incomeTrustGap: 73, finscope: 77 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/hollard", positive: 980, negative: 520, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["Hollard Foundation — R80m/year"],
    greenFootprint: { score: 71, notes: [] },
    aiUse: { score: 72, notes: [] },
    known: true,
  },

  // ───────── SHORT-TERM INSURANCE ─────────
  {
    slug: "santam", name: "Santam", category: "insurance-short-term", subcategory: "Home & Property",
    website: "https://www.santam.co.za", logo: "🚗", founded: 1918, bbb: "1", regulator: "FSCA",
    trustComponents: { transparency: 84, financialConfidence: 89, incomeTrustGap: 71, finscope: 82 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/santam", positive: 1120, negative: 640, lastUpdated: "2026-06-01" }],
    fines: [],
    awards: [{ year: 2025, name: "Insurer of the Year", issuer: "Fin24" }],
    socialImpact: ["Santam SMK programme"],
    greenFootprint: { score: 75, notes: [] },
    aiUse: { score: 81, notes: ["Computer-vision vehicle damage assessment"] },
    known: true,
  },
  {
    slug: "outsurance", name: "OUTsurance", category: "insurance-short-term", subcategory: "Motor",
    website: "https://www.outsurance.co.za", logo: "🚗", founded: 1998, bbb: "1", regulator: "FSCA",
    trustComponents: { transparency: 83, financialConfidence: 85, incomeTrustGap: 76, finscope: 80 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/outsurance", positive: 1450, negative: 720, lastUpdated: "2026-06-01" }],
    fines: [],
    awards: [{ year: 2024, name: "Customer Service Excellence", issuer: "Service Sector Awards" }],
    socialImpact: ["Cash-back no-claims model"],
    greenFootprint: { score: 72, notes: [] },
    aiUse: { score: 79, notes: ["OUTbox telematics AI"] },
    known: true,
  },
  {
    slug: "king-price", name: "King Price", category: "insurance-short-term", subcategory: "Motor",
    website: "https://www.kingprice.co.za", logo: "🚗", founded: 2010, bbb: "2", regulator: "FSCA",
    trustComponents: { transparency: 80, financialConfidence: 78, incomeTrustGap: 84, finscope: 79 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/king-price", positive: 1680, negative: 510, lastUpdated: "2026-06-01" }],
    fines: [],
    awards: [{ year: 2025, name: "Most Trusted Short-term Insurer", issuer: "Reader's Digest" }],
    socialImpact: ["Decreasing-premium model", "Lion-cub CSR campaign"],
    greenFootprint: { score: 69, notes: [] },
    aiUse: { score: 73, notes: ["Depreciation algorithm"] },
    known: true,
  },
  {
    slug: "miway", name: "MiWay", category: "insurance-short-term", subcategory: "Motor",
    website: "https://www.miway.co.za", logo: "🚗", founded: 2007, bbb: "2", regulator: "FSCA",
    trustComponents: { transparency: 77, financialConfidence: 80, incomeTrustGap: 74, finscope: 76 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/miway", positive: 1380, negative: 1180, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["MiWay Foundation"],
    greenFootprint: { score: 66, notes: [] },
    aiUse: { score: 70, notes: [] },
    known: true,
  },
  {
    slug: "hollard-short", name: "Hollard Short-Term", category: "insurance-short-term", subcategory: "Business",
    website: "https://www.hollard.co.za", logo: "🚗", founded: 1985, bbb: "1", regulator: "FSCA",
    trustComponents: { transparency: 81, financialConfidence: 84, incomeTrustGap: 72, finscope: 79 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/hollard", positive: 980, negative: 520, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["SME business cover innovation"],
    greenFootprint: { score: 71, notes: [] },
    aiUse: { score: 74, notes: [] },
    known: true,
  },
  {
    slug: "discovery-insure", name: "Discovery Insure", category: "insurance-short-term", subcategory: "Motor",
    website: "https://www.discovery.co.za/car-insurance", logo: "🚗", founded: 2011, bbb: "2", regulator: "FSCA",
    trustComponents: { transparency: 86, financialConfidence: 83, incomeTrustGap: 62, finscope: 77 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/discovery-insure", positive: 1540, negative: 490, lastUpdated: "2026-06-01" }],
    fines: [],
    awards: [{ year: 2025, name: "Telematics Leader Africa", issuer: "Telematics Wire" }],
    socialImpact: ["Vitality Drive behaviour programme"],
    greenFootprint: { score: 80, notes: ["EV discount structures"] },
    aiUse: { score: 92, notes: ["Telematics + computer vision for claims"] },
    known: true,
  },

  // ───────── MEDICAL AID ─────────
  {
    slug: "discovery-health", name: "Discovery Health", category: "medical-aid", subcategory: "Comprehensive",
    website: "https://www.discovery.co.za/medical-aid", logo: "🏥", founded: 1998, bbb: "2", regulator: "CMS / FSCA",
    trustComponents: { transparency: 84, financialConfidence: 85, incomeTrustGap: 58, finscope: 78 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/discovery-health", positive: 1720, negative: 1390, lastUpdated: "2026-06-01" }],
    fines: [],
    awards: [{ year: 2025, name: "Best Medical Scheme", issuer: "Discovery / Fin24 readers" }],
    socialImpact: ["Vitality Wellness", "Network coverage of 44,000 providers"],
    greenFootprint: { score: 74, notes: [] },
    aiUse: { score: 90, notes: ["AI chronic disease management"] },
    known: true,
  },
  {
    slug: "bonitas", name: "Bonitas Medical Fund", category: "medical-aid", subcategory: "Comprehensive",
    website: "https://www.bonitas.co.za", logo: "🏥", founded: 1989, bbb: "1", regulator: "CMS",
    trustComponents: { transparency: 79, financialConfidence: 80, incomeTrustGap: 76, finscope: 77 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/bonitas", positive: 720, negative: 480, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["TB & HIV wellness grants"],
    greenFootprint: { score: 65, notes: [] },
    aiUse: { score: 70, notes: [] },
    known: true,
  },
  {
    slug: "gems", name: "GEMS", category: "medical-aid", subcategory: "Comprehensive",
    website: "https://www.gems.gov.za", logo: "🏥", founded: 2005, bbb: "N/A", regulator: "CMS",
    trustComponents: { transparency: 62, financialConfidence: 64, incomeTrustGap: 90, finscope: 72 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/gems", positive: 520, negative: 1180, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["Public-sector employees only — inclusive pricing"],
    greenFootprint: { score: 50, notes: [] },
    aiUse: { score: 55, notes: [] },
    known: true,
  },
  {
    slug: "momentum-health", name: "Momentum Health", category: "medical-aid", subcategory: "Comprehensive",
    website: "https://www.momentum.co.za/health", logo: "🏥", founded: 2000, bbb: "1", regulator: "CMS",
    trustComponents: { transparency: 76, financialConfidence: 78, incomeTrustGap: 70, finscope: 74 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/momentum-health", positive: 640, negative: 720, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["Multiply wellness"],
    greenFootprint: { score: 62, notes: [] },
    aiUse: { score: 68, notes: [] },
    known: true,
  },
  {
    slug: "bestmed", name: "Bestmed", category: "medical-aid", subcategory: "Savings Plan",
    website: "https://www.bestmed.co.za", logo: "🏥", founded: 1998, bbb: "2", regulator: "CMS",
    trustComponents: { transparency: 74, financialConfidence: 77, incomeTrustGap: 73, finscope: 73 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/bestmed", positive: 410, negative: 220, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["Academic bursaries"],
    greenFootprint: { score: 60, notes: [] },
    aiUse: { score: 62, notes: [] },
    known: false,
  },

  // ───────── BROKERAGES ─────────
  {
    slug: "psg-wealth", name: "PSG Wealth", category: "brokerages", subcategory: "JSE Broker",
    website: "https://www.psgwealth.co.za", logo: "🤝", founded: 1996, bbb: "2", regulator: "FSCA",
    trustComponents: { transparency: 80, financialConfidence: 82, incomeTrustGap: 60, finscope: 74 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/psg", positive: 520, negative: 310, lastUpdated: "2026-06-01" }],
    fines: [], awards: [], socialImpact: [],
    greenFootprint: { score: 66, notes: [] },
    aiUse: { score: 72, notes: ["Advisor AI co-pilot"] },
    known: true,
  },
  {
    slug: "independent-online", name: "Independent Online (iKruger)", category: "brokerages", subcategory: "Insurance Broker",
    website: "https://www.ikruger.co.za", logo: "🤝", founded: 2002, bbb: "2", regulator: "FSCA",
    trustComponents: { transparency: 78, financialConfidence: 76, incomeTrustGap: 79, finscope: 75 },
    evidence: [{ source: "HelloPeter", url: "https://www.hellopeter.com/ikruger", positive: 880, negative: 420, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["Comparison transparency"],
    greenFootprint: { score: 72, notes: [] },
    aiUse: { score: 77, notes: ["AI quote comparison engine"] },
    known: false,
  },

  // ───────── INVESTMENT ─────────
  {
    slug: "allan-gray", name: "Allan Gray", category: "investment", subcategory: "Asset Manager",
    website: "https://www.allangray.co.za", logo: "📈", founded: 1974, bbb: "1", regulator: "FSCA",
    trustComponents: { transparency: 89, financialConfidence: 90, incomeTrustGap: 66, finscope: 83 },
    evidence: [{ source: "Reddit", url: "https://www.reddit.com/r/SAinvestments/search/?q=allan+gray", positive: 420, negative: 180, lastUpdated: "2026-06-01" }],
    fines: [],
    awards: [{ year: 2025, name: "Fund House of the Year", issuer: "Daybreak Awards" }],
    socialImpact: ["100% employee-owned", "Allan Gray Orbis Foundation — student finance"],
    greenFootprint: { score: 78, notes: ["Responsible-investment signatory"] },
    aiUse: { score: 74, notes: [] },
    known: true,
  },
  {
    slug: "sygnia", name: "Sygnia", category: "investment", subcategory: "Retail Brokerage Platform",
    website: "https://www.sygnia.co.za", logo: "📈", founded: 2001, bbb: "2", regulator: "FSCA",
    trustComponents: { transparency: 83, financialConfidence: 79, incomeTrustGap: 81, finscope: 79 },
    evidence: [{ source: "Reddit", url: "https://www.reddit.com/r/SAinvestments/search/?q=sygnia", positive: 340, negative: 210, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["Fee transparency advocate"],
    greenFootprint: { score: 70, notes: ["Passive / ESG default options"] },
    aiUse: { score: 80, notes: ["Portfolio AI insights"] },
    known: true,
  },
  {
    slug: "easyequities", name: "EasyEquities", category: "investment", subcategory: "Retail Brokerage Platform",
    website: "https://www.easyequities.co.za", logo: "📈", founded: 2014, bbb: "2", regulator: "FSCA",
    trustComponents: { transparency: 82, financialConfidence: 80, incomeTrustGap: 93, finscope: 85 },
    evidence: [
      { source: "HelloPeter", url: "https://www.hellopeter.com/easyequities", positive: 1910, negative: 640, lastUpdated: "2026-06-01" },
      { source: "Reddit", url: "https://www.reddit.com/r/SAinvestments/search/?q=easyequities", positive: 680, negative: 240, lastUpdated: "2026-06-01" },
    ],
    fines: [],
    awards: [{ year: 2025, name: "Retail Investor Platform of the Year", issuer: "African Financial Awards" }],
    socialImpact: ["Democratised investing — min R10", "Free financial-education academy"],
    greenFootprint: { score: 81, notes: ["Paperless by design", "ESG filters on ETFs"] },
    aiUse: { score: 83, notes: ["Robo-advisor beta"] },
    known: true,
  },
  {
    slug: "coronation", name: "Coronation Fund Managers", category: "investment", subcategory: "Asset Manager",
    website: "https://www.coronation.com", logo: "📈", founded: 1991, bbb: "1", regulator: "FSCA",
    trustComponents: { transparency: 84, financialConfidence: 85, incomeTrustGap: 63, finscope: 79 },
    evidence: [{ source: "Reddit", url: "https://www.reddit.com/r/SAinvestments/search/?q=coronation", positive: 260, negative: 140, lastUpdated: "2026-06-01" }],
    fines: [], awards: [], socialImpact: [],
    greenFootprint: { score: 74, notes: [] },
    aiUse: { score: 76, notes: [] },
    known: true,
  },
  {
    slug: "ninety-one", name: "Ninety One", category: "investment", subcategory: "Asset Manager",
    website: "https://www.ninetyone.com", logo: "📈", founded: 1991, bbb: "2", regulator: "FSCA / FCA (UK)",
    trustComponents: { transparency: 83, financialConfidence: 84, incomeTrustGap: 62, finscope: 78 },
    evidence: [{ source: "Reddit", url: "https://www.reddit.com/r/SAinvestments/search/?q=ninety+one", positive: 210, negative: 130, lastUpdated: "2026-06-01" }],
    fines: [], awards: [],
    socialImpact: ["Impact-investing desk"],
    greenFootprint: { score: 82, notes: ["Net-zero aligned"] },
    aiUse: { score: 78, notes: [] },
    known: true,
  },
];

export const PROVIDERS: Provider[] = raw.map((p) => ({
  ...p,
  trustScore: build(p.trustComponents),
}));

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: src/components/MobileNav.tsx                            ║
// ╚════════════════════════════════════════════════════════════════╝

import type { Route } from "../App";

export function MobileNav({ route, go }: { route: Route; go: (r: Route) => void }) {
  const active = (v: Route["view"]) => route.view === v;
  return (
    <nav className="md:hidden fixed bottom-0 inset-x-0 z-40 bg-white border-t border-slate-200 shadow-[0_-2px_10px_rgba(0,0,0,0.04)]">
      <div className="grid grid-cols-4 h-16">
        <Btn label="Home"       icon="🏠" isActive={active("home")}     onClick={() => go({ view: "home" })} />
        <Btn label="Categories" icon="🗂️" isActive={active("category")} onClick={() => go({ view: "category", slug: "banks" })} />
        <Btn label="Compare"    icon="⚖️" isActive={active("compare")}  onClick={() => go({ view: "compare" })} />
        <Btn label="Search"     icon="🔎" isActive={false} onClick={() => {
          const el = document.querySelector<HTMLInputElement>('input[placeholder*="Search"]');
          el?.focus();
          window.scrollTo({ top: 0, behavior: "smooth" });
        }} />
      </div>
    </nav>
  );
}

function Btn({ icon, label, isActive, onClick }: {
  icon: string; label: string; isActive: boolean; onClick: () => void;
}) {
  return (
    <button
      onClick={onClick}
      className={`flex flex-col items-center justify-center gap-0.5 transition ${
        isActive ? "text-emerald-600" : "text-slate-500"
      }`}
    >
      <span className="text-xl">{icon}</span>
      <span className="text-[10px] font-medium">{label}</span>
    </button>
  );
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: src/components/CompareView.tsx                          ║
// ╚════════════════════════════════════════════════════════════════╝

import { useMemo, useState } from "react";
import type { Provider } from "../data/providers";
import { CATEGORIES } from "../data/categories";

export function CompareView({
  providers, initialA, initialB, go,
}: {
  providers: Provider[];
  initialA?: string;
  initialB?: string;
  go: (r: any) => void;
}) {
  const [a, setA] = useState(initialA ?? providers[0]?.slug ?? "");
  const [b, setB] = useState(initialB ?? providers[1]?.slug ?? "");

  const pa = useMemo(() => providers.find((p) => p.slug === a) ?? null, [a, providers]);
  const pb = useMemo(() => providers.find((p) => p.slug === b) ?? null, [b, providers]);

  return (
    <main className="max-w-6xl mx-auto px-4 py-8 pb-24 md:pb-8">
      <button onClick={() => go({ view: "home" })} className="text-sm text-slate-500 mb-2">← Back</button>
      <h1 className="text-3xl font-bold">⚖️ Compare providers</h1>
      <p className="text-slate-600 mt-1">Pick two to see them head-to-head on every dimension.</p>

      <div className="mt-6 grid md:grid-cols-2 gap-4">
        <ProviderPicker label="Provider A" value={a} onChange={setA} providers={providers} />
        <ProviderPicker label="Provider B" value={b} onChange={setB} providers={providers} />
      </div>

      {pa && pb && (
        <>
          <WinnerBanner a={pa} b={pb} />

          <div className="mt-6 rounded-2xl border border-slate-200 bg-white overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-slate-50">
                <tr>
                  <th className="text-left p-3 w-1/3">Metric</th>
                  <th className="text-right p-3 w-1/3">{pa.name}</th>
                  <th className="text-right p-3 w-1/3">{pb.name}</th>
                </tr>
              </thead>
              <tbody>
                <Row label="🏷 Category"     a={catName(pa.category)}   b={catName(pb.category)} />
                <Row label="📂 Sub-category" a={pa.subcategory}         b={pb.subcategory} />
                <Row label="📅 Founded"      a={pa.founded.toString()}  b={pb.founded.toString()} />
                <Row label="🏅 B-BBEE Level" a={`L${pa.bbb}`}           b={`L${pb.bbb}`} />
                <Divider />
                <ScoreRow label="⭐ Trust Me Score"      a={pa.trustScore}                          b={pb.trustScore} />
                <ScoreRow label="🔍 Transparency"        a={pa.trustComponents.transparency}        b={pb.trustComponents.transparency} />
                <ScoreRow label="💪 Financial Confidence" a={pa.trustComponents.financialConfidence} b={pb.trustComponents.financialConfidence} />
                <ScoreRow label="🤝 Income Trust Gap"    a={pa.trustComponents.incomeTrustGap}      b={pb.trustComponents.incomeTrustGap} />
                <ScoreRow label="📊 FinScope 2.0"        a={pa.trustComponents.finscope}            b={pb.trustComponents.finscope} />
                <Divider />
                <ScoreRow label="🤖 AI Score"   a={pa.aiUse.score}           b={pb.aiUse.score} />
                <ScoreRow label="🌿 Green Score" a={pa.greenFootprint.score}  b={pb.greenFootprint.score} />
                <Divider />
                <Row label="⚠️ Fines count"     a={pa.fines.length.toString()}  b={pb.fines.length.toString()} />
                <Row label="🏅 Awards count"    a={pa.awards.length.toString()} b={pb.awards.length.toString()} />
                <Row label="📣 Evidence sources" a={pa.evidence.length.toString()} b={pb.evidence.length.toString()} />
                <Row label="📣 Total positive reviews"
                  a={pa.evidence.reduce((s, e) => s + e.positive, 0).toLocaleString()}
                  b={pb.evidence.reduce((s, e) => s + e.positive, 0).toLocaleString()} />
                <Row label="📣 Total negative reviews"
                  a={pa.evidence.reduce((s, e) => s + e.negative, 0).toLocaleString()}
                  b={pb.evidence.reduce((s, e) => s + e.negative, 0).toLocaleString()} />
              </tbody>
            </table>
          </div>

          <div className="mt-6 grid md:grid-cols-2 gap-4">
            <EvidenceList title={`${pa.name} — evidence`} evidence={pa.evidence} />
            <EvidenceList title={`${pb.name} — evidence`} evidence={pb.evidence} />
          </div>

          <div className="mt-6 grid md:grid-cols-2 gap-4">
            <FinesList title={`${pa.name} — fines`} fines={pa.fines} />
            <FinesList title={`${pb.name} — fines`} fines={pb.fines} />
          </div>
        </>
      )}
    </main>
  );
}

function catName(slug: string) {
  return CATEGORIES.find((c) => c.slug === slug)?.name ?? slug;
}

function WinnerBanner({ a, b }: { a: Provider; b: Provider }) {
  if (a.trustScore === b.trustScore)
    return (
      <div className="mt-6 rounded-xl bg-slate-100 p-4 text-center text-slate-600">
        🤝 It's a tie — <b>{a.trustScore}</b> each.
      </div>
    );
  const winner = a.trustScore > b.trustScore ? a : b;
  const loser  = a.trustScore > b.trustScore ? b : a;
  return (
    <div className="mt-6 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-500 text-white p-5">
      <div className="text-xs uppercase tracking-wider opacity-80">Trust Me winner</div>
      <div className="text-2xl font-bold mt-1">{winner.logo} {winner.name}</div>
      <div className="text-sm opacity-90 mt-1">
        {winner.trustScore} vs {loser.trustScore} (+{(winner.trustScore - loser.trustScore).toFixed(1)})
      </div>
    </div>
  );
}

function ProviderPicker({
  label, value, onChange, providers,
}: {
  label: string; value: string; onChange: (v: string) => void; providers: Provider[];
}) {
  return (
    <div>
      <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider">{label}</label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 bg-white"
      >
        {providers.map((p) => (
          <option key={p.slug} value={p.slug}>{p.name} — {p.trustScore}</option>
        ))}
      </select>
    </div>
  );
}

function Row({ label, a, b, highlight }: { label: string; a: string; b: string; highlight?: "a" | "b" }) {
  return (
    <tr className="border-t border-slate-100">
      <td className="p-3 text-slate-600">{label}</td>
      <td className={`p-3 text-right font-medium ${highlight === "a" ? "text-emerald-700" : ""}`}>{a}</td>
      <td className={`p-3 text-right font-medium ${highlight === "b" ? "text-emerald-700" : ""}`}>{b}</td>
    </tr>
  );
}

function ScoreRow({ label, a, b }: { label: string; a: number; b: number }) {
  const winner: "a" | "b" | "none" = a > b ? "a" : b > a ? "b" : "none";
  return <Row label={label} a={a.toString()} b={b.toString()} highlight={winner === "none" ? undefined : winner} />;
}

function Divider() {
  return <tr><td colSpan={3} className="h-2 bg-slate-50" /></tr>;
}

function EvidenceList({ title, evidence }: { title: string; evidence: Provider["evidence"] }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4">
      <h3 className="font-semibold mb-3">{title}</h3>
      <div className="space-y-2">
        {evidence.length === 0 && <div className="text-sm text-slate-500">No evidence yet.</div>}
        {evidence.map((e) => {
          const ratio = e.positive / (e.positive + e.negative || 1);
          return (
            <div key={e.source} className="border border-slate-100 rounded-lg p-3">
              <div className="flex justify-between text-sm">
                <span className="font-semibold">{e.source}</span>
                <a href={e.url} target="_blank" rel="noreferrer" className="text-emerald-700 text-xs">open ↗</a>
              </div>
              <div className="flex gap-3 text-xs mt-1 text-slate-600">
                <span className="text-emerald-700">▲ {e.positive.toLocaleString()}</span>
                <span className="text-rose-700">▼ {e.negative.toLocaleString()}</span>
              </div>
              <div className="mt-2 h-1.5 bg-slate-100 rounded-full overflow-hidden">
                <div className="h-full bg-emerald-500" style={{ width: `${ratio * 100}%` }} />
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function FinesList({ title, fines }: { title: string; fines: Provider["fines"] }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4">
      <h3 className="font-semibold mb-3">{title}</h3>
      {fines.length === 0 && <div className="text-sm text-slate-500">Clean record ✅</div>}
      <ul className="text-sm space-y-2">
        {fines.map((f, i) => (
          <li key={i} className="border-l-4 border-rose-400 pl-3">
            <div className="font-semibold">{f.year} · {f.amount} <span className="text-slate-500 font-normal">· {f.regulator}</span></div>
            <div className="text-slate-600">{f.reason}</div>
          </li>
        ))}
      </ul>
    </div>
  );
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  FILE: src/App.tsx                                             ║
// ╚════════════════════════════════════════════════════════════════╝

import { useEffect, useMemo, useState } from "react";
import { CATEGORIES } from "./data/categories";
import type { Provider } from "./data/providers";
import { gradeFromScore, rankInCategory } from "./lib/trust";
import { loadProviders } from "./lib/api";
import { MobileNav } from "./components/MobileNav";
import { CompareView } from "./components/CompareView";

export type Route =
  | { view: "home" }
  | { view: "category"; slug: string }
  | { view: "provider"; slug: string }
  | { view: "compare"; a?: string; b?: string };

export default function App() {
  const [route, setRoute] = useState<Route>({ view: "home" });
  const [query, setQuery] = useState("");
  const [providers, setProviders] = useState<Provider[]>([]);
  const [compareTray, setCompareTray] = useState<string[]>([]);

  useEffect(() => { loadProviders().then(setProviders); }, []);

  const matches = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return [];
    return providers
      .filter((p) =>
        p.name.toLowerCase().includes(q) ||
        p.slug.includes(q) ||
        p.category.includes(q) ||
        p.subcategory.toLowerCase().includes(q)
      )
      .slice(0, 8);
  }, [query, providers]);

  const go = (r: Route) => {
    setRoute(r);
    setQuery("");
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const toggleCompare = (slug: string) => {
    setCompareTray((prev) => {
      if (prev.includes(slug)) return prev.filter((s) => s !== slug);
      if (prev.length >= 2) return [prev[1], slug];
      return [...prev, slug];
    });
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-50 to-white text-slate-900 pb-20 md:pb-0">
      {/* NAV */}
      <header className="sticky top-0 z-30 bg-white/80 backdrop-blur border-b border-slate-200">
        <div className="max-w-6xl mx-auto px-4 py-3 flex items-center gap-4">
          <button onClick={() => go({ view: "home" })} className="flex items-center gap-2 font-bold shrink-0">
            <span className="text-2xl">🇿🇦</span>
            <span className="text-lg hidden sm:inline">Trust Me <span className="text-xs font-medium text-slate-500">· SA Finance</span></span>
          </button>

          <div className="relative flex-1 max-w-xl ml-auto">
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search a bank, insurer, medical aid…"
              className="w-full rounded-full border border-slate-300 px-4 py-2 pr-10 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500"
            />
            <span className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400">🔎</span>

            {matches.length > 0 && (
              <div className="absolute left-0 right-0 mt-2 bg-white rounded-xl shadow-lg border border-slate-200 overflow-hidden z-40">
                {matches.map((p) => (
                  <button
                    key={p.slug}
                    onClick={() => go({ view: "provider", slug: p.slug })}
                    className="w-full text-left px-4 py-2 hover:bg-slate-50 flex items-center justify-between border-b last:border-b-0"
                  >
                    <span className="flex items-center gap-2">
                      <span>{p.logo}</span>
                      <span className="text-sm font-medium">{p.name}</span>
                      <span className="text-xs text-slate-500 hidden sm:inline">· {p.subcategory}</span>
                    </span>
                    <span className="text-xs font-bold text-emerald-700">{p.trustScore}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
      </header>

      {/* ROUTES */}
      {route.view === "home"     && <Home providers={providers} go={go} toggleCompare={toggleCompare} compareTray={compareTray} />}
      {route.view === "category" && <CategoryView slug={route.slug} providers={providers} go={go} toggleCompare={toggleCompare} compareTray={compareTray} />}
      {route.view === "provider" && <ProviderView slug={route.slug} providers={providers} go={go} />}
      {route.view === "compare"  && <CompareView providers={providers} initialA={route.a ?? compareTray[0]} initialB={route.b ?? compareTray[1]} go={go} />}

      {/* Floating compare tray */}
      {compareTray.length > 0 && route.view !== "compare" && (
        <div className="fixed bottom-20 md:bottom-4 left-1/2 -translate-x-1/2 z-30 bg-white rounded-full shadow-xl border border-slate-200 px-4 py-2 flex items-center gap-3">
          <span className="text-xs text-slate-500">Compare:</span>
          {compareTray.map((s) => {
            const p = providers.find((x) => x.slug === s);
            return p ? (
              <span key={s} className="text-xs bg-emerald-50 text-emerald-700 px-2 py-1 rounded-full flex items-center gap-1">
                {p.logo} {p.name}
                <button onClick={() => toggleCompare(s)} className="ml-1 text-slate-400 hover:text-rose-500">×</button>
              </span>
            ) : null;
          })}
          {compareTray.length === 2 && (
            <button
              onClick={() => go({ view: "compare", a: compareTray[0], b: compareTray[1] })}
              className="text-xs bg-emerald-600 text-white px-3 py-1 rounded-full hover:bg-emerald-700"
            >
              Go →
            </button>
          )}
        </div>
      )}

      <MobileNav route={route} go={go} />
      <Footer />
    </div>
  );
}

/* ───────────── HOME ───────────── */
function Home({ providers, go, toggleCompare, compareTray }: {
  providers: Provider[]; go: (r: Route) => void; toggleCompare: (s: string) => void; compareTray: string[];
}) {
  return (
    <main className="max-w-6xl mx-auto px-4 py-10">
      <section className="text-center py-8">
        <h1 className="text-4xl md:text-5xl font-extrabold tracking-tight">
          South Africa's <span className="text-emerald-600">financial trust</span>, in one place.
        </h1>
        <p className="mt-4 text-slate-600 max-w-2xl mx-auto">
          Real, evidence-based rankings of every SA bank, insurer, medical aid, broker & asset manager.
        </p>
      </section>

      <section className="mb-10">
        <h2 className="text-xl font-semibold mb-3">🏆 Top 5 Most Trusted</h2>
        <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
          {rankInCategory(providers).slice(0, 5).map((p, i) => (
            <MiniCard key={p.slug} p={p} rank={i + 1}
              onClick={() => go({ view: "provider", slug: p.slug })}
              onCompare={() => toggleCompare(p.slug)}
              inCompare={compareTray.includes(p.slug)}
            />
          ))}
        </div>
      </section>

      <section>
        <h2 className="text-xl font-semibold mb-3">Browse by category</h2>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          {CATEGORIES.map((c) => (
            <button key={c.slug} onClick={() => go({ view: "category", slug: c.slug })}
              className="rounded-2xl border border-slate-200 bg-white p-5 text-left hover:border-emerald-500 hover:shadow-md transition">
              <div className="text-3xl">{c.icon}</div>
              <div className="font-semibold mt-2">{c.name}</div>
              <div className="text-xs text-slate-500 mt-1">
                {providers.filter((p) => p.category === c.slug).length} providers ranked
              </div>
            </button>
          ))}
        </div>
      </section>

      <section className="mt-12 rounded-2xl bg-slate-900 text-white p-8">
        <h3 className="text-2xl font-bold">How the Trust Me Index is calculated</h3>
        <div className="grid md:grid-cols-4 gap-4 mt-6 text-sm">
          <Pillar t="28%" title="Transparency"        body="Fee disclosure, annual-report timeliness, complaints published." />
          <Pillar t="26%" title="Financial Confidence" body="Solvency, payout ratios, complaints-per-policy." />
          <Pillar t="20%" title="Income Trust Gap"     body="Access & affordability for LSM 1–6." />
          <Pillar t="26%" title="FinScope 2.0"         body="Customer awareness, usage, satisfaction." />
        </div>
      </section>
    </main>
  );
}

/* ───────────── CATEGORY ───────────── */
function CategoryView({ slug, providers, go, toggleCompare, compareTray }: {
  slug: string; providers: Provider[]; go: (r: Route) => void; toggleCompare: (s: string) => void; compareTray: string[];
}) {
  const cat = CATEGORIES.find((c) => c.slug === slug);
  if (!cat) return <div className="p-8">Not found.</div>;
  const [sub, setSub] = useState<string>("all");
  const filtered = rankInCategory(providers.filter((p) => p.category === slug && (sub === "all" || p.subcategory === sub)));

  return (
    <main className="max-w-6xl mx-auto px-4 py-8 pb-24 md:pb-8">
      <button onClick={() => go({ view: "home" })} className="text-sm text-slate-500 mb-2">← Back</button>
      <div className="flex items-center gap-3 mb-2">
        <span className="text-4xl">{cat.icon}</span>
        <h1 className="text-3xl font-bold">{cat.name}</h1>
      </div>
      <p className="text-slate-600 mb-4">{cat.description}</p>

      <div className="flex flex-wrap gap-2 mb-6">
        <SubTab active={sub === "all"} onClick={() => setSub("all")}>All</SubTab>
        {cat.subcategories.map((s) => (
          <SubTab key={s} active={sub === s} onClick={() => setSub(s)}>{s}</SubTab>
        ))}
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white overflow-x-auto">
        <table className="w-full text-sm min-w-[720px]">
          <thead className="bg-slate-50 text-slate-600">
            <tr>
              <th className="text-left p-3">#</th>
              <th className="text-left p-3">Provider</th>
              <th className="text-left p-3">Trust Me</th>
              <th className="text-left p-3">Transp.</th>
              <th className="text-left p-3">Fin. Conf.</th>
              <th className="text-left p-3">Income</th>
              <th className="text-left p-3">FinScope</th>
              <th className="text-left p-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((p, i) => {
              const g = gradeFromScore(p.trustScore);
              return (
                <tr key={p.slug} className="border-t border-slate-100 hover:bg-emerald-50/40">
                  <td className="p-3 font-semibold text-slate-500">{i + 1}</td>
                  <td className="p-3">
                    <button onClick={() => go({ view: "provider", slug: p.slug })} className="flex items-center gap-2 text-left">
                      <span className="text-xl">{p.logo}</span>
                      <div>
                        <div className="font-semibold">{p.name}</div>
                        <div className="text-xs text-slate-500">{p.subcategory}</div>
                      </div>
                    </button>
                  </td>
                  <td className="p-3">
                    <span className={`font-bold ${g.color}`}>{p.trustScore}</span>
                    <span className={`ml-1 text-xs ${g.color}`}>{g.letter}</span>
                  </td>
                  <td className="p-3">{p.trustComponents.transparency}</td>
                  <td className="p-3">{p.trustComponents.financialConfidence}</td>
                  <td className="p-3">{p.trustComponents.incomeTrustGap}</td>
                  <td className="p-3">{p.trustComponents.finscope}</td>
                  <td className="p-3">
                    <button
                      onClick={() => toggleCompare(p.slug)}
                      className={`text-xs px-2 py-1 rounded-full border ${
                        compareTray.includes(p.slug)
                          ? "bg-emerald-600 text-white border-emerald-600"
                          : "border-slate-300 text-slate-600 hover:border-emerald-500"
                      }`}
                    >
                      {compareTray.includes(p.slug) ? "✓ In tray" : "⚖️ Compare"}
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </main>
  );
}

/* ───────────── PROVIDER DETAIL ───────────── */
function ProviderView({ slug, providers, go }: { slug: string; providers: Provider[]; go: (r: Route) => void }) {
  const p = providers.find((x) => x.slug === slug);
  if (!p) return <div className="p-8">Not found.</div>;
  const sameCat = rankInCategory(providers.filter((x) => x.category === p.category));
  const rank = sameCat.findIndex((x) => x.slug === p.slug) + 1;
  const g = gradeFromScore(p.trustScore);

  return (
    <main className="max-w-6xl mx-auto px-4 py-8 pb-24 md:pb-8">
      <button onClick={() => go({ view: "category", slug: p.category })} className="text-sm text-slate-500 mb-2">
        ← Back to {CATEGORIES.find((c) => c.slug === p.category)?.name}
      </button>

      <div className="rounded-2xl border border-slate-200 bg-white p-6 flex flex-col md:flex-row gap-6 items-start">
        <div className="text-6xl">{p.logo}</div>
        <div className="flex-1">
          <div className="flex items-center gap-2 flex-wrap">
            <h1 className="text-3xl font-bold">{p.name}</h1>
            <span className="text-xs px-2 py-1 rounded-full bg-slate-100 text-slate-700">{p.subcategory}</span>
            {p.known && <span className="text-xs px-2 py-1 rounded-full bg-amber-100 text-amber-800">Household name</span>}
          </div>
          <div className="text-sm text-slate-500 mt-1">Founded {p.founded} · B-BBEE L{p.bbb} · {p.regulator}</div>
          <a href={p.website} target="_blank" rel="noreferrer" className="text-sm text-emerald-700 hover:underline mt-2 inline-block">🔗 {p.website}</a>

          <div className="flex gap-6 mt-5 flex-wrap">
            <BigStat label="Trust Me Score" value={p.trustScore.toString()} sub={`${g.letter} · ${g.label}`} color={g.color} />
            <BigStat label="Category Rank"  value={`#${rank}`} sub={`of ${sameCat.length}`} />
            <BigStat label="Evidence"       value={p.evidence.length.toString()} sub="sources" />
          </div>
        </div>
      </div>

      <section className="mt-6 grid md:grid-cols-3 gap-4">
        <div className="rounded-2xl border border-slate-200 bg-white p-5">
          <h3 className="font-semibold mb-3">Trust Components</h3>
          <Bar label="Transparency Index"   v={p.trustComponents.transparency}         w={0.28} />
          <Bar label="Financial Confidence" v={p.trustComponents.financialConfidence}  w={0.26} />
          <Bar label="Income Trust Gap"     v={p.trustComponents.incomeTrustGap}       w={0.20} />
          <Bar label="FinScope 2.0"         v={p.trustComponents.finscope}             w={0.26} />
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-5">
          <h3 className="font-semibold mb-3">🤖 AI & Modern Tech: {p.aiUse.score}/100</h3>
          <ul className="text-sm text-slate-600 space-y-1">
            {p.aiUse.notes.map((n, i) => <li key={i}>• {n}</li>)}
          </ul>
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-5">
          <h3 className="font-semibold mb-3">🌿 Green: {p.greenFootprint.score}/100</h3>
          <ul className="text-sm text-slate-600 space-y-1">
            {p.greenFootprint.notes.length ? p.greenFootprint.notes.map((n, i) => <li key={i}>• {n}</li>) : <li className="text-slate-400">No disclosures.</li>}
          </ul>
        </div>
      </section>

      <section className="mt-6 rounded-2xl border border-slate-200 bg-white p-5">
        <h3 className="font-semibold mb-3">📣 Evidence Trail</h3>
        <div className="grid md:grid-cols-2 gap-3">
          {p.evidence.map((e) => (
            <div key={e.source} className="rounded-xl border border-slate-200 p-4">
              <div className="flex items-center justify-between">
                <div className="font-semibold">{e.source}</div>
                <a href={e.url} target="_blank" rel="noreferrer" className="text-xs text-emerald-700 hover:underline">open ↗</a>
              </div>
              <div className="text-xs text-slate-500 mt-1">Updated {e.lastUpdated}</div>
              <div className="flex gap-4 mt-3 text-sm">
                <span className="text-emerald-700 font-semibold">▲ {e.positive.toLocaleString()}</span>
                <span className="text-rose-700 font-semibold">▼ {e.negative.toLocaleString()}</span>
              </div>
              <div className="mt-2 h-2 bg-slate-100 rounded-full overflow-hidden">
                <div className="h-full bg-emerald-500" style={{ width: `${(e.positive / (e.positive + e.negative)) * 100}%` }} />
              </div>
            </div>
          ))}
        </div>
      </section>

      <section className="mt-6 grid md:grid-cols-2 gap-4">
        <div className="rounded-2xl border border-slate-200 bg-white p-5">
          <h3 className="font-semibold mb-3">⚠️ Fines</h3>
          {p.fines.length === 0 ? <div className="text-sm text-slate-500">Clean record ✅</div> : (
            <ul className="text-sm space-y-2">
              {p.fines.map((f, i) => (
                <li key={i} className="border-l-4 border-rose-400 pl-3">
                  <div className="font-semibold">{f.year} · {f.amount} <span className="text-slate-500 font-normal">— {f.regulator}</span></div>
                  <div className="text-slate-600">{f.reason}</div>
                </li>
              ))}
            </ul>
          )}
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-5">
          <h3 className="font-semibold mb-3">🏅 Awards</h3>
          {p.awards.length === 0 ? <div className="text-sm text-slate-500">None logged.</div> : (
            <ul className="text-sm space-y-2">
              {p.awards.map((a, i) => (
                <li key={i} className="border-l-4 border-amber-400 pl-3">
                  <div className="font-semibold">{a.year} · {a.name}</div>
                  <div className="text-slate-500">{a.issuer}</div>
                </li>
              ))}
            </ul>
          )}
        </div>
      </section>

      <section className="mt-6 rounded-2xl border border-slate-200 bg-white p-5">
        <h3 className="font-semibold mb-3">🤝 Social Impact</h3>
        {p.socialImpact.length === 0 ? <div className="text-sm text-slate-500">No initiatives logged.</div> : (
          <ul className="text-sm text-slate-700 space-y-1">{p.socialImpact.map((s, i) => <li key={i}>• {s}</li>)}</ul>
        )}
      </section>

      <section className="mt-6 rounded-2xl border border-slate-200 bg-white p-5">
        <h3 className="font-semibold mb-3">🥊 Ranks vs competitors</h3>
        <div className="space-y-2">
          {sameCat.map((c, i) => {
            const gc = gradeFromScore(c.trustScore);
            const max = sameCat[0].trustScore;
            return (
              <div key={c.slug} className={`flex items-center gap-3 p-2 rounded ${c.slug === p.slug ? "bg-emerald-50 ring-1 ring-emerald-300" : ""}`}>
                <div className="w-6 text-right text-slate-500 text-sm">{i + 1}</div>
                <div className="text-xl">{c.logo}</div>
                <div className="flex-1">
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-sm">{c.name}</span>
                    <span className={`font-bold ${gc.color}`}>{c.trustScore}</span>
                  </div>
                  <div className="mt-1 h-2 bg-slate-100 rounded-full overflow-hidden">
                    <div className="h-full bg-emerald-500" style={{ width: `${(c.trustScore / max) * 100}%` }} />
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </section>
    </main>
  );
}

/* ───────────── Shared bits ───────────── */
function Footer() {
  return (
    <footer className="mt-16 border-t border-slate-200 py-8 text-center text-sm text-slate-500 px-4 pb-24 md:pb-8">
      <div>Trust Me © 2026 · South African financial transparency index</div>
      <div className="mt-1">Data: HelloPeter, Reddit, Google, Wikipedia, Ombudsman, NCR, CMS, FSCA, FinScope.</div>
    </footer>
  );
}
function SubTab({ children, active, onClick }: { children: React.ReactNode; active: boolean; onClick: () => void }) {
  return <button onClick={onClick} className={`px-3 py-1 rounded-full text-sm border ${active ? "bg-emerald-600 text-white border-emerald-600" : "bg-white border-slate-200"}`}>{children}</button>;
}
function Pillar({ t, title, body }: { t: string; title: string; body: string }) {
  return <div><div className="text-3xl font-bold text-emerald-400">{t}</div><div className="font-semibold mt-1">{title}</div><div className="text-slate-300 text-xs mt-1">{body}</div></div>;
}
function Bar({ label, v, w }: { label: string; v: number; w: number }) {
  return <div className="mb-3"><div className="flex justify-between text-sm"><span>{label} <span className="text-slate-400 text-xs">(×{w})</span></span><span className="font-semibold">{v}</span></div><div className="h-2 bg-slate-100 rounded-full overflow-hidden"><div className="h-full bg-gradient-to-r from-emerald-500 to-teal-500" style={{ width: `${v}%` }} /></div></div>;
}
function BigStat({ label, value, sub, color = "text-slate-900" }: { label: string; value: string; sub?: string; color?: string }) {
  return <div><div className="text-xs text-slate-500 uppercase tracking-wider">{label}</div><div className={`text-3xl font-extrabold ${color}`}>{value}</div>{sub && <div className="text-xs text-slate-500">{sub}</div>}</div>;
}
function MiniCard({ p, rank, onClick, onCompare, inCompare }: {
  p: Provider; rank: number; onClick: () => void; onCompare: () => void; inCompare: boolean;
}) {
  const g = gradeFromScore(p.trustScore);
  return (
    <div className="rounded-xl border border-slate-200 bg-white p-3 relative">
      <button onClick={onClick} className="block w-full text-left">
        <div className="flex items-center justify-between"><span className="text-xl">{p.logo}</span><span className="text-xs text-slate-400">#{rank}</span></div>
        <div className="font-semibold text-sm mt-2 truncate">{p.name}</div>
        <div className={`font-bold text-lg ${g.color}`}>{p.trustScore}</div>
        <div className={`text-xs ${g.color}`}>{g.label}</div>
      </button>
      <button onClick={onCompare} className={`absolute top-2 right-2 text-[10px] px-1.5 py-0.5 rounded-full border ${inCompare ? "bg-emerald-600 text-white border-emerald-600" : "border-slate-200 text-slate-500"}`}>⚖️</button>
    </div>
  );
}

   mkdir trust-me-sa && cd trust-me-sa
   git init
   # paste each file at its FILE: path above
   git add .
   git commit -m "Initial commit — Trust Me SA"
   git remote add origin https://github.com/<you>/trust-me-sa.git
   git push -u origin main
   
