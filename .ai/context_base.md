# Claw Client Assessment — AI Context

## Project Overview
Client needs assessment toolkit for AI assistant deployments. Two-part questionnaire (Private + Enterprise) focused on understanding client needs in plain language — no technical jargon. Maps needs internally to Claw platforms (OpenClaw, NanoClaw, PicoClaw), LLM models, and skills.

## Repository Structure
```
claw-client-assessment/
├── .ai/
│   └── context_base.md              # This file — AI agent context
├── questionnaire/
│   ├── client-intake-form.json       # Machine-readable questionnaire schema
│   └── needs-mapping-matrix.json     # Internal: needs → platform/model/skills
├── benchmarks/
│   ├── platform-comparison.json      # Internal: OpenClaw vs NanoClaw vs PicoClaw
│   ├── llm-model-comparison.json     # Internal: LLM pricing, benchmarks
│   └── skills-catalog.json           # Internal: top skills by category
├── packages/
│   └── service-packages.json         # Pricing and service tier definitions
├── scripts/
│   └── generate_questionnaire.py     # Python script to regenerate the DOCX
├── docs/
│   └── AI_Agent_Client_Needs_Assessment.docx
├── LICENSE
└── README.md
```

## Questionnaire Structure (Client-Facing DOCX)
- **Part A (Private):** A1 About You, A2 Digital Life, A3 Wishlist, A4 Capabilities (40 checkboxes), A5 Integration, A6 Privacy
- **Part B (Enterprise):** B1 Company Profile, B2 Pain Points, B3 Capabilities (40 checkboxes), B4 Integration, B5 Compliance, B6 Scale
- **Section C:** Pricing with cost estimation by daily usage
- **Section D:** Authorization and sign-off

## Pricing (EUR)
- Private: €1,000 one-time
- Enterprise: From €5,000 one-time
- Managed Service: €300/month (installation included)
- Ongoing Assistance: €500/month (after 6 months)
- API costs: on the client (€5–€200/month depending on usage)
- Hosting: client's own hardware, their cloud, or we provide infrastructure

## Key Data Points
- 40 capability checkboxes in human-friendly language per audience
- 15 needs-to-solution mappings (internal reference)
- 9 LLM models with pricing (internal reference)
- 3 platforms compared across 16+ dimensions (internal reference)
- Cost estimation table: Light/Moderate/Heavy/Intensive usage tiers

## Related Repos
- https://github.com/Amenthyx/claw-one-click-deploy — Deployment scripts and configs
- https://github.com/Amenthyx/amenthyx-ai-teams — Amenthyx AI Teams framework
