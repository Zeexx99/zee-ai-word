/**
 * Built-in report templates. Inserted as Markdown → Word headings (Heading 1/2), lists and tables,
 * so References → Table of Contents works straight away. Text in [brackets] is a placeholder.
 */

export interface ReportTemplate {
  id: string;
  name: string;
  description: string;
  outline: string;
}

export const TEMPLATES: ReportTemplate[] = [
  {
    id: "daily-production",
    name: "Daily production report",
    description: "Field/well rates, downtime, ESP status and actions for one day.",
    outline: `# Daily Production Report – [Field] – [Date]

## 1. Summary
[Two or three sentences: total oil, water and gas, change vs yesterday, main events.]

## 2. Production by well
| Well | Oil (bbl/d) | Water (bbl/d) | Gas (Mscf/d) | WC (%) | Status |
| --- | --- | --- | --- | --- | --- |
| [W-1] | [ ] | [ ] | [ ] | [ ] | [On/Off] |

## 3. Downtime and deferred production
- [Well] – [hours] h – [cause] – [deferred bbl]

## 4. ESP / artificial lift status
- [Well]: frequency [Hz], intake pressure [psi], motor temperature [°F], current [A] – [remarks]

## 5. HSE
- [Incidents, near misses or "No HSE events"]

## 6. Actions and next 24 hours
| Action | Owner | Due |
| --- | --- | --- |
| [ ] | [ ] | [ ] |`,
  },
  {
    id: "esp-failure",
    name: "ESP failure analysis (RCA)",
    description: "Root-cause analysis of an ESP failure with run life, teardown and actions.",
    outline: `# ESP Failure Analysis – [Well] – [Date]

## 1. Executive summary
[What failed, run life, most likely root cause and the key recommendation.]

## 2. Well and equipment data
| Item | Value |
| --- | --- |
| Well / field | [ ] |
| Pump / stages | [ ] |
| Motor (HP / V / A) | [ ] |
| Install date / failure date | [ ] |
| Run life (days) | [ ] |

## 3. Operating history before failure
[Rates, frequency, intake pressure, motor temperature and current trends; trips and restarts.]

## 4. Pull and teardown findings
- [Component] – [observation] – [photo reference]

## 5. Root-cause analysis
### 5.1 Possible causes considered
- [Cause] – [evidence for / against]
### 5.2 Most likely root cause
[Statement with supporting evidence.]

## 6. Recommendations
| Recommendation | Owner | Due |
| --- | --- | --- |
| [ ] | [ ] | [ ] |

## 7. Lessons learned
- [ ]`,
  },
  {
    id: "well-test",
    name: "Well test report",
    description: "Test conditions, stabilised rates, PI/IPR and conclusions.",
    outline: `# Well Test Report – [Well] – [Date]

## 1. Objective
[Why the test was run.]

## 2. Test conditions
| Parameter | Value |
| --- | --- |
| Choke size | [ ] |
| Separator pressure / temperature | [ ] |
| Test duration | [ ] |

## 3. Stabilised results
| Parameter | Value | Unit |
| --- | --- | --- |
| Oil rate | [ ] | bbl/d |
| Water rate | [ ] | bbl/d |
| Gas rate | [ ] | Mscf/d |
| Water cut | [ ] | % |
| GOR | [ ] | scf/STB |
| FBHP / SBHP | [ ] | psi |

## 4. Analysis
[Productivity index, IPR (e.g. Vogel), comparison with the previous test.]

## 5. Conclusions and recommendations
- [ ]`,
  },
  {
    id: "moc",
    name: "Management of Change (MOC)",
    description: "Proposed change, justification, risk assessment and approvals.",
    outline: `# Management of Change – [Title]

## 1. Description of the change
[What will change, where and for how long (temporary / permanent).]

## 2. Justification
[Why the change is needed; expected benefit.]

## 3. Risk assessment
| Hazard | Consequence | Likelihood | Risk | Mitigation |
| --- | --- | --- | --- | --- |
| [ ] | [ ] | [ ] | [ ] | [ ] |

## 4. Impact on procedures, drawings and training
- [P&IDs, operating procedures, training needs]

## 5. Implementation plan
| Step | Owner | Date |
| --- | --- | --- |
| [ ] | [ ] | [ ] |

## 6. Approvals
| Role | Name | Signature | Date |
| --- | --- | --- | --- |
| Originator | [ ] | | |
| Technical authority | [ ] | | |
| HSE | [ ] | | |`,
  },
  {
    id: "tech-memo",
    name: "Technical memo",
    description: "Short engineering memo: background, analysis, recommendation.",
    outline: `# Technical Memo – [Subject]

**To:** [ ]  
**From:** [ ]  
**Date:** [ ]

## 1. Purpose
[One paragraph.]

## 2. Background
[Context and data used.]

## 3. Analysis
[Method, results, assumptions.]

## 4. Recommendation
[Clear recommendation and next steps.]`,
  },
  {
    id: "minutes",
    name: "Meeting minutes",
    description: "Attendees, decisions and action items.",
    outline: `# Meeting Minutes – [Topic] – [Date]

**Attendees:** [ ]

## 1. Agenda
1. [ ]

## 2. Discussion
- [ ]

## 3. Decisions
- [ ]

## 4. Action items
| Action | Owner | Due |
| --- | --- | --- |
| [ ] | [ ] | [ ] |`,
  },
];
