import React, { useState, useRef } from 'react';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun } from 'docx';
import { saveAs } from 'file-saver';
import { marked } from 'marked';
import katex from 'katex';
import { toPng } from 'html-to-image';
import { motion, AnimatePresence } from 'framer-motion';
import { FileDown, ClipboardPaste, Sparkles, CheckCircle2, Calculator } from 'lucide-react';
import 'katex/dist/katex.min.css'; // Import KaTeX styles

// --- Helper: Latex to Image Generator ---
// We render the math to a hidden DOM node, snapshot it, then remove it.
const generateMathImage = async (latex) => {
  const node = document.createElement('div');
  // Styling to make sure the image looks crisp in Word
  Object.assign(node.style, {
    position: 'absolute',
    top: '-9999px',
    left: '-9999px',
    padding: '20px',
    backgroundColor: 'white',
    fontSize: '24px', // Larger font = better resolution in Word
    display: 'inline-block' 
  });
  document.body.appendChild(node);

  try {
    // Render Latex into the node
    katex.render(latex, node, {
      throwOnError: false,
      displayMode: true, // Center large equations
    });

    // Convert to PNG Blob
    const dataUrl = await toPng(node, { quality: 1.0, pixelRatio: 3 }); // High pixel ratio for sharpness
    const res = await fetch(dataUrl);
    return await res.blob();
  } catch (e) {
    console.error("Math render error", e);
    return null;
  } finally {
    document.body.removeChild(node);
  }
};

const App = () => {
  const [text, setText] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [success, setSuccess] = useState(false);
  const [progress, setProgress] = useState('');

  // --- LOGIC: The Master Converter ---
  const generateDocx = async () => {
    if (!text.trim()) return;
    setIsGenerating(true);
    setProgress('Parsing text...');

    // 1. Split text into segments: [Text, Math, Text, Math...]
    // Regex matches $$...$$ (block) or \(...\) (inline)
    const regex = /(\$\$[\s\S]*?\$\$|\\\[[\s\S]*?\\\]|\\\([\s\S]*?\\\))/g;
    const parts = text.split(regex);

    const docChildren = [];

    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      if (!part.trim()) continue;

      // Check if this part is Math
      const isBlockMath = part.startsWith('$$') || part.startsWith('\\[');
      const isInlineMath = part.startsWith('\\(');

      if (isBlockMath || isInlineMath) {
        // Clean the tags (remove $$ or \())
        const cleanLatex = part
          .replace(/^\$\$|\$\$?$/g, '')
          .replace(/^\\\[|\\\]$/g, '')
          .replace(/^\\\(|\\\)$/g, '');

        setProgress(`Rendering Equation ${i + 1}...`);
        const imageBlob = await generateMathImage(cleanLatex);

        if (imageBlob) {
          docChildren.push(
            new Paragraph({
              children: [
                new ImageRun({
                  data: imageBlob,
                  transformation: { width: 300, height: 100 }, // Aspect ratio is auto-handled usually, but we set max width
                }),
              ],
              spacing: { after: 200, before: 200 },
              alignment: "center", // Center equations
            })
          );
        }
      } else {
        // It's Markdown Text -> Use Marked to process
        const tokens = marked.lexer(part);
        tokens.forEach(token => {
          if (token.type === 'heading') {
            const levelMap = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
            docChildren.push(new Paragraph({
              text: token.text,
              heading: levelMap[token.depth] || HeadingLevel.HEADING_1,
              spacing: { after: 120, before: 240 },
            }));
          } else if (token.type === 'list') {
            token.items.forEach(item => {
               docChildren.push(new Paragraph({
                 text: item.text,
                 bullet: { level: 0 }
               }));
            });
          } else {
             // Standard Paragraph
             docChildren.push(new Paragraph({
               text: token.text || token.raw,
               spacing: { after: 120 },
             }));
          }
        });
      }
    }

    setProgress('Finalizing .docx...');
    
    // Create Document
    const doc = new Document({
      sections: [{ children: docChildren }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "AI_Export_With_Math.docx");

    setIsGenerating(false);
    setSuccess(true);
    setTimeout(() => setSuccess(false), 3000);
  };

  const handlePaste = async () => {
    try {
      const t = await navigator.clipboard.readText();
      setText(t);
    } catch (e) { alert("Please paste manually."); }
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] text-slate-800 font-sans flex flex-col items-center py-12 px-4">
      
      {/* Header */}
      <motion.div initial={{ opacity: 0, y: -20 }} animate={{ opacity: 1, y: 0 }} className="text-center mb-8">
        <h1 className="text-4xl font-extrabold tracking-tight text-slate-900 flex items-center justify-center gap-3">
          <span className="bg-blue-600 text-white p-2 rounded-lg"><Calculator size={28}/></span>
          AI to Word
        </h1>
        <p className="mt-3 text-slate-500 text-lg">Export GPT/Gemini chats with <b>Equations</b> perfectly preserved.</p>
      </motion.div>

      {/* Main Interface */}
      <motion.div 
        initial={{ opacity: 0, scale: 0.98 }}
        animate={{ opacity: 1, scale: 1 }}
        className="w-full max-w-4xl bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden flex flex-col md:flex-row h-[600px]"
      >
        
        {/* Input Side */}
        <div className="flex-1 flex flex-col border-r border-slate-100">
          <div className="p-4 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
            <span className="font-semibold text-slate-600 text-sm">Paste AI Text Here</span>
            <button onClick={handlePaste} className="text-blue-600 hover:text-blue-700 text-sm font-medium flex items-center gap-1">
              <ClipboardPaste size={14} /> Paste
            </button>
          </div>
          <textarea 
            className="flex-1 p-6 resize-none focus:outline-none text-slate-700 leading-relaxed font-mono text-sm"
            placeholder="Paste text with $$ equations here..."
            value={text}
            onChange={(e) => setText(e.target.value)}
          />
        </div>

        {/* Preview / Action Side */}
        <div className="md:w-72 bg-slate-50 p-6 flex flex-col justify-center items-center text-center gap-6">
           <div className="w-full">
             <h3 className="font-semibold text-slate-700 mb-2">Ready to Export?</h3>
             <p className="text-xs text-slate-400">Supports Markdown, Lists, and $$ LaTeX $$</p>
           </div>

           <motion.button
             whileHover={{ scale: 1.05 }}
             whileTap={{ scale: 0.95 }}
             onClick={generateDocx}
             disabled={!text || isGenerating}
             className={`w-full py-4 rounded-xl font-bold text-white shadow-lg flex flex-col items-center justify-center gap-2 transition-all
               ${!text ? 'bg-slate-300 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-500 shadow-blue-500/30'}`}
           >
             {isGenerating ? (
               <>
                 <span className="animate-spin rounded-full h-6 w-6 border-b-2 border-white"></span>
                 <span className="text-xs font-normal opacity-90">{progress}</span>
               </>
             ) : success ? (
               <> <CheckCircle2 size={28} /> Saved! </>
             ) : (
               <> <FileDown size={28} /> Convert Now </>
             )}
           </motion.button>

           <div className="text-[10px] text-slate-400 max-w-[200px]">
             Equations are rendered as high-res images to ensure compatibility with all Word versions.
           </div>
        </div>

      </motion.div>
    </div>
  );
};

export default App;