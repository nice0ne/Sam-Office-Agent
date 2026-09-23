# Word Corporate Styler & Contract/Compliance Reviewer Design Spec

## 1. Overview & Business Value

In daily enterprise and corporate office work, Microsoft Word is heavily used for:
1. **Contracts, Service Level Agreements (SLAs), Non-Disclosure Agreements (NDAs), and Vendor Agreements**.
2. **Standard Operating Procedures (SOPs), Formal Memorandums, and Board Papers**.

Two major pain points exist:
- **Contract Review Friction**: Business managers and legal teams spend significant time reading contracts to check if key protections (liability caps, SLA penalties, termination terms, payment deadlines) are present, favorable, or risky.
- **Inconsistent Document Styling**: Corporate documents often suffer from mismatched fonts, unformatted headings, irregular line spacing, and inconsistent bullet formatting when multiple team members contribute.

This feature adds:
1. **Contract & Compliance Clause Reviewer (`review_compliance_clauses`)**: Scans contract or policy text to detect key clauses, evaluate risks (unlimited liability, ambiguous deadlines, penalty imbalances), and generate structured audit findings with actionable clause-level recommendations.
2. **Corporate Brand & Document Styler (`apply_corporate_style`)**: Applies standardized enterprise typography, heading styles, and paragraph spacing across the document or selection.
3. **Quick Action Presets in Chat UI**: Direct one-click access for `⚖️ Review Kontrak & Risiko` and `🎨 Format Brand Korporat`.

---

## 2. Architecture & Data Contracts

### 2.1 Type Definitions (`src/services/office/types.ts`)

```typescript
export interface ComplianceClause {
  category: 'payment_terms' | 'sla_performance' | 'liability_indemnity' | 'termination_cancellation' | 'confidentiality_nda' | 'dispute_resolution' | 'force_majeure' | 'general';
  excerpt: string;
  status: 'compliant' | 'warning' | 'high_risk' | 'missing';
  analysis: string;
  recommendation?: string;
}

export interface ComplianceReviewOptions {
  scope?: 'selection' | 'document';
  contractType?: 'vendor_service' | 'employment' | 'nda' | 'procurement' | 'general';
  strictness?: 'standard' | 'strict';
}

export interface ComplianceReviewResult {
  contractType: string;
  overallRiskLevel: 'low' | 'medium' | 'high';
  clausesReviewedCount: number;
  identifiedClauses: ComplianceClause[];
  missingCriticalClauses: string[];
  executiveSummary: string;
  actionableRecommendations: string[];
}

export interface CorporateStyleOptions {
  theme?: 'corporate_navy' | 'executive_emerald' | 'modern_minimalist' | 'official_government';
  scope?: 'selection' | 'document';
  fontFamily?: string;
  applyHeadingHierarchy?: boolean;
}

export interface CorporateStyleResult {
  appliedTheme: string;
  fontFamily: string;
  styledParagraphsCount: number;
  headingsCount: number;
  message: string;
}
```

### 2.2 Driver Interface Extensions (`IOfficeDriver` & `IDocumentDriver`)

```typescript
export interface IDocumentDriver extends IOfficeDriver {
  // Existing methods...
  reviewComplianceClauses?(options?: ComplianceReviewOptions): Promise<ComplianceReviewResult>;
  applyCorporateStyle?(options?: CorporateStyleOptions): Promise<CorporateStyleResult>;
}
```

---

## 3. Implementation Details

### 3.1 Contract & Compliance Review Engine
- In `mockDriver.ts`:
  - Scans mock text or document outline.
  - Matches regex and heuristics for:
    - SLA & penalties (e.g. `denda`, `penalti`, `ganti rugi`, `SLA`, `terlambat`, `uptime`).
    - Termination (e.g. `pemutusan`, `pengakhiran`, `30 hari`, `sepihak`).
    - Confidentiality (e.g. `rahasia`, `NDA`, `informasi rahasia`).
    - Liability limits (e.g. `tanggung jawab`, `liability`, `unlimited`, `tidak terbatas`).
  - Evaluates risk levels: flags unlimited liability or missing force majeure as `high_risk` / `warning`.
  - Produces structured `ComplianceReviewResult`.
- In `wordDriver.ts`:
  - Fetches text from selection or full body using `context.document.getSelection()` or `context.document.body`.
  - Analyzes the content and generates findings.

### 3.2 Corporate Brand Styler
- In `mockDriver.ts`:
  - Simulates styling paragraphs and headings, returning count and applied theme details.
- In `wordDriver.ts`:
  - Uses Office.js Word batching:
    - Loads `context.document.body.paragraphs`.
    - Detects headings based on text patterns or existing styles.
    - Applies font family (e.g. "Calibri" / "Segoe UI" / "Aptos" / "Times New Roman").
    - Formats Title/H1 (accent color, bold, spacing) and H2/Body (standard line spacing, clean margins).
    - Executes within `await context.sync()`.

### 3.3 Word Specialist Agent (`src/agents/word/wordAgent.ts`)
- Registers `review_compliance_clauses` and `apply_corporate_style` in `getTools()`.
- Implements execution and returns clean, professional Markdown reports with risk badges (`🟢 LOW`, `🟡 MEDIUM`, `🔴 HIGH RISK`).

### 3.4 Quick Action Presets (`src/components/Chat/QuickActionPresets.tsx`)
- Adds:
  - `⚖️ Review Kontrak & Risiko`
  - `🎨 Format Brand Korporat`
