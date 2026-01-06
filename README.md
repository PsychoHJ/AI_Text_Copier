# AI to Word (DOCX) Converter

A fast, privacy-focused web tool that converts AI chat exports (from ChatGPT, Gemini, Claude) into perfectly formatted Microsoft Word documents. 

Built to look like **BibCit**, featuring a clean utility-focused interface, instant processing, and robust support for mathematical equations.

![Project Preview](https://via.placeholder.com/800x400?text=App+Screenshot+Here)

## 🚀 Features

* **Markdown Support:** Preserves bolding, lists, and headers from AI responses.
* **LaTeX Equation Support:** Automatically renders complex math (`$$...$$` or `\[...\]`) as high-quality images inside Word.
* **Privacy First:** Zero backend. All processing happens entirely in your browser using `docx.js`.
* **Modern UI:** Clean, animated interface built with Tailwind CSS and Framer Motion.
* **Auto Paste:** One-click clipboard detection.

## 🛠️ Tech Stack

* **Framework:** React + Vite
* **Styling:** Tailwind CSS
* **Animations:** Framer Motion
* **Word Generation:** docx.js
* **Math Rendering:** KaTeX + html-to-image

## 📦 Installation & Setup

1.  **Clone the repository**
    ```bash
    git clone [https://github.com/your-username/ai-to-docx.git](https://github.com/your-username/ai-to-docx.git)
    cd ai-to-docx
    ```

2.  **Install dependencies**
    ```bash
    npm install
    ```

3.  **Run the development server**
    ```bash
    npm run dev
    ```

## 🚢 Deployment

This project is optimized for deployment on **Vercel** or **Netlify**.
Simply connect this repository to your Vercel dashboard for automatic deployment.