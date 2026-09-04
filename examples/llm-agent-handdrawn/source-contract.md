# LLM agent hand-drawn overview source contract

Status: brief-driven example; no target overview is being recreated

## Source authority

- Communication brief: explain one inference-time LLM-agent loop in a single editable PowerPoint slide.
- Scientific scope: generic architecture, not a claim about one vendor, benchmark, or model family.
- Semantic authorities:
  - the Transformer pattern for token-conditioned generation;
  - retrieval as evidence added to model context;
  - an agent/controller executing a model-proposed tool call and returning the observation to context.
- Evidence scope: explanatory schematic only; no empirical image, score, accuracy, latency, or safety claim.
- Deliverables: editable PowerPoint source, PowerPoint-rendered PNG, design specification, and fresh audit report.

## Figure Claim

An LLM agent assembles a question and retrieved evidence into context, uses the LLM to propose either an answer or a tool request, executes requested tools through a controller, returns observations to context, and produces a grounded answer with visible evidence links.

## Required-node ledger

| id | exact label | role | priority |
| --- | --- | --- | --- |
| n_question | Question | user input | primary |
| n_sources | Retrieved sources | external evidence | primary |
| n_context | Context builder | retrieval, tokenization, assembly | focal-support |
| n_tokens | Context tokens | model input sequence | primary |
| n_llm | LLM core | attention and feed-forward processing | focal |
| n_next | next-token / tool-call proposal | model output decision | primary |
| n_controller | Agent controller | executes the proposed action | primary |
| n_tool | External tool | calculator, search, or API | secondary |
| n_observation | Observation | tool result returned to context | primary |
| n_answer | Grounded answer | user-facing output | focal-support |
| n_citations | Evidence links | source-grounding indicator | secondary |

## Required-edge ledger

| id | source → target | visible relation |
| --- | --- | --- |
| e_question_context | Question → Context builder | query enters context |
| e_sources_context | Retrieved sources → Context builder | evidence is assembled into context |
| e_context_tokens | Context builder → Context tokens | produces ordered token sequence |
| e_tokens_llm | Context tokens → LLM core | conditions generation |
| e_llm_next | LLM core → next-token / tool-call proposal | proposes continuation or action |
| e_next_answer | proposal → Grounded answer | answer branch |
| e_next_controller | proposal → Agent controller | action branch |
| e_controller_tool | Agent controller → External tool | executes the tool call |
| e_tool_observation | External tool → Observation | returns tool result |
| e_observation_context | Observation → Context builder | one explicit feedback lane |
| e_sources_citations | Retrieved sources → Evidence links | provenance retained |
| e_citations_answer | Evidence links → Grounded answer | grounding is visible in output |

## Negative-path checks

- Retrieved sources do not bypass the context builder to become generated prose.
- A tool observation does not update model weights during ordinary inference.
- The LLM proposes a tool call; the controller executes it.
- Evidence links do not imply that next-token prediction itself verifies truth.
- The observation loop is optional and bounded, not an infinite decorative cycle.

## Style-source firewall

These references influence visual language only; none supplies this example's labels, topology, or composition.

| source | transferable traits | prohibited transfer |
| --- | --- | --- |
| [Vicky, “Not a Lecture from the Man on the Hill” (2026)](https://medium.com/@vicky01010110/not-a-lecture-from-the-man-on-the-hill-61bcb6c4c66b) | monochrome technical doodles, generous whitespace, short hand-lettered labels | its LLM layout and wording |
| [Karina Lewis, “Context windows — 5 perspectives on gen AI” (2025)](https://medium.com/@karinasketchesthings/context-windows-5-perspectives-on-gen-ai-c9dddb9fe45e) | restrained green-style highlighting, region rhythm, human sketchnote voice | the five-topic structure and metaphors |
| [Maarten Grootendorst, “A Visual Guide to LLM Agents” (2025)](https://newsletter.maartengrootendorst.com/p/a-visual-guide-to-llm-agents) | modular visual grammar and clean semantic chunking | exact diagrams, icons, and composition |
| [Henrik Kniberg, “Generative AI in a Nutshell” (2024)](https://www.youtube.com/watch?v=2IK3DFHRFfw) | selective highlighter marks and compact explanatory doodles | poster-scale density and source-specific drawings |

## Acceptance contract

- All 11 required nodes and 12 required relations are recoverable from the latest render.
- The five negative paths are not visually implied.
- The hand-drawn style is visible beyond the title font, while all semantic connectors remain exact.
- The slide contains no picture object and no copied reference asset.
- The native PowerPoint object graph and the PowerPoint renderer both pass review.
