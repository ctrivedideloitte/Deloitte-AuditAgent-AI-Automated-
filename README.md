# Audit Agent 🕵️‍♂️

A powerful, AI-driven reconciliation tool designed to streamline the training audit process. This application leverages the Gemini 3 Flash model to extract data from source reports (even PDFs!) and cross-reference them against internal directories to generate high-precision "nudge" communications.

## 🚀 Key Features

- **Automated Data Extraction**: Upload training reports in XLSX, CSV, or PDF format.
- **PDF-to-Excel Intelligence**: Integrated AI-powered conversion that transforms unstructured PDFs into structured Excel data for better audit accuracy.
- **Intelligent Matching**: Cross-references "File A" (Source) against "File B" (Directory) using fuzzy logic to find missing emails and contact details.
- **Outlook Integration**: Automatically generates and dispatches professional reconciliation drafts directly to Microsoft Outlook.
- **Session Persistence**: Saves your audits and CC memory locally so you never lose your progress.
- **Optimized UI**: A sleek, Deloitte-inspired bento grid interface with collapsible sidebars and real-time processing feedback.

## 🛠 Tech Stack

- **Frontend**: React 18+ with Vite
- **Styling**: Tailwind CSS
- **Animations**: Framer Motion (motion/react)
- **AI Engine**: Google Gemini 1.5 Flash
- **Data Processing**: SheetJS (XLSX)
- **Icons**: Lucide React

## 📦 Getting Started

### Prerequisites

- Node.js 18+
- npm or yarn

### Installation

1. Clone the repository:
   ```bash
   git clone <your-repo-url>
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Create a `.env` file based on `.env.example` and add your Gemini API Key:
   ```env
   VITE_GEMINI_API_KEY=your_api_key_here
   ```
4. Start the development server:
   ```bash
   npm run dev
   ```

## 🔒 Data Privacy

This tool is designed for professional audit contexts. Ensure you adhere to regional privacy protocols when handling internal employee data.

---
Built with ❤️ using Google AI Studio.
