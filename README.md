# IXL Curriculum Emulation & Extraction Engine

A resilient, single-tab browser automation engine built with Python and Playwright designed to systematically map, navigate, and extract mathematical sequence variations from interactive educational platforms.

##  The Engineering Challenge
Most conventional web scrapers collapse when interacting with modern educational platforms due to strict server-side defenses. This project was developed to overcome three primary systemic bottlenecks:
1. **State-Driven Single Page Applications (SPAs):** The target platform updates question interfaces dynamically via JavaScript without refreshing the DOM layout.
2. **Behavioral Anti-Bot Telemetry:** Rapid URL hopping or unauthenticated deep-linking triggers instant session death and hard redirects.
3. **Layout Composition Variance:** Question elements shape-shift continuously between standard text blocks, embedded vector code graphics (SVGs), and rich accessible layers.

## 🛠️ System Architecture & Solutions Implemented
* **Human-Emulation Pipeline:** Relies entirely on manual navigation behaviors, forcing Playwright to interact with physical sidebar elements and breadcrumbs to preserve referral chains naturally.
* **Component-Root Text Flattening:** Uses a custom JavaScript extraction script to clone dynamic DOM states, physically stripping interactive answer choices in memory to isolate clean question strings.
* **Synchronized State Checking:** Features a polling loop that tracks system telemetry values to block the extraction sequencer until new layouts are confirmed hydrated.
* **Visual Asset Harvesting:** Dynamically assesses layouts for graphical containers (Number lines, coordinate planes, counting matrices) to take perfectly cropped localized snapshots.

##  System Architecture Workflow
1. **Authentication Handshake:** Log in, bake localized cookie states, and preserve connection tokens globally.
2. **Curriculum Mapping:** Scans index layouts to map hierarchical categories seamlessly.
3. **Execution Sequence:** Clicks into target skill modules -> Executes extraction -> Solves multi-tier confirmation popups -> Advances state -> Uses breadcrumbs to return safely without session drops.

##  Installation & Usage

1. Clone the repository:
   ```bash
   git clone [https://github.com/shub404/IXL-Emulate-Engine.git](https://github.com/shub404/IXL-Emulate-Engine.git)
   cd IXL-Emulate-Engine