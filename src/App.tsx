/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useRef } from 'react';
import JSZip from 'jszip';
import { Upload, FileDown, ShieldCheck, Eraser, Loader2, AlertCircle, Trash2, CheckCircle2 } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// --- Types ---

interface FileState {
  file: File;
  name: string;
  size: string;
}

// --- Utilities ---

const formatBytes = (bytes: number) => {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

export default function App() {
  const [fileState, setFileState] = useState<FileState | null>(null);
  const [watermark, setWatermark] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [isDone, setIsDone] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [processedFile, setProcessedFile] = useState<Blob | null>(null);
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (!file.name.endsWith('.pptx')) {
        setError('仅支持 .pptx 格式的文件');
        return;
      }
      setFileState({
        file,
        name: file.name,
        size: formatBytes(file.size)
      });
      setError(null);
      setIsDone(false);
      setProcessedFile(null);
    }
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) {
      if (!file.name.endsWith('.pptx')) {
        setError('仅支持 .pptx 格式的文件');
        return;
      }
      setFileState({
        file,
        name: file.name,
        size: formatBytes(file.size)
      });
      setError(null);
      setIsDone(false);
      setProcessedFile(null);
    }
  }, []);

  const [removedCount, setRemovedCount] = useState(0);

  const removeWatermark = async () => {
    if (!fileState || !watermark) return;

    setIsProcessing(true);
    setError(null);
    setRemovedCount(0);
    let totalRemoved = 0;

    try {
      const zip = new JSZip();
      const content = await zip.loadAsync(fileState.file);
      
      const parser = new DOMParser();
      const serializer = new XMLSerializer();

      // We will iterate over ALL files in the ppt directory
      for (const [path, file] of Object.entries(content.files)) {
        if (path.endsWith('.xml') || path.endsWith('.rels')) {
          let xmlString = await file.async('string');
          let fileModified = false;

          // Strategy 1: Aggressive Cross-Tag Regex replacement
          // This handles text split across many nodes, and removes the text content
          // It's a fallback that works even if DOM structure is weird
          const escapedWatermark = watermark.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const crossTagPattern = escapedWatermark.split('').join('(?:<[^>]+>)*');
          const regex = new RegExp(crossTagPattern, 'gi');
          
          if (regex.test(xmlString)) {
            const matches = xmlString.match(regex);
            if (matches) totalRemoved += matches.length;
            xmlString = xmlString.replace(regex, '');
            fileModified = true;
          }

          // Strategy 2: DOM-based Shape/Group removal
          // This targets the container (to catch icons/logos next to text)
          if (path.endsWith('.xml')) {
            const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
            let domModified = false;
            const elementsToRemove = new Set<Element>();

            const allNodes = xmlDoc.getElementsByTagNameNS('*', '*');
            for (let i = 0; i < allNodes.length; i++) {
              const node = allNodes[i];
              const textContent = node.textContent || '';
              const altText = node.getAttribute('descr') || node.getAttribute('title') || node.getAttribute('name') || '';

              const isMatch = (str: string) => 
                str.toLowerCase().includes(watermark.toLowerCase()) || 
                str.toLowerCase().includes('notebooklm');

              if (isMatch(textContent) || isMatch(altText)) {
                // Find parent shape to delete the whole group
                let current: Node | null = node;
                while (current && current.parentNode) {
                  const nodeName = (current as Element).localName || '';
                  if (['sp', 'pic', 'grpSp'].includes(nodeName)) {
                    elementsToRemove.add(current as Element);
                    break;
                  }
                  current = current.parentNode;
                }
              }
            }

            elementsToRemove.forEach(el => {
              if (el.parentNode) {
                el.parentNode.removeChild(el);
                domModified = true;
                totalRemoved++;
              }
            });

            if (domModified) {
              xmlString = serializer.serializeToString(xmlDoc);
              fileModified = true;
            }
          }

          if (fileModified) {
            zip.file(path, xmlString);
          }
        }
      }

      if (totalRemoved === 0) {
        setError(`未找到水印 "${watermark}"。请检查文字是否完全匹配，或者尝试只输入关键词（如 "Notebook"）。`);
        setIsProcessing(false);
        return;
      }

      const generatedContent = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
      setProcessedFile(generatedContent);
      setRemovedCount(totalRemoved);
      setIsDone(true);
    } catch (err) {
      console.error(err);
      setError('处理失败：PPT 文件可能已加密或损坏。');
    } finally {
      setIsProcessing(false);
    }
  };

  const downloadFile = () => {
    if (!processedFile || !fileState) return;
    const url = URL.createObjectURL(processedFile);
    const a = document.createElement('a');
    a.href = url;
    a.download = `cleaned_${fileState.name}`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const reset = () => {
    setFileState(null);
    setWatermark('');
    setIsDone(false);
    setProcessedFile(null);
    setError(null);
    setRemovedCount(0);
  };


  return (
    <div className="min-h-screen bg-[#FDFCFB] text-slate-900 font-sans selection:bg-orange-100">
      {/* Header */}
      <nav className="border-b border-slate-100 bg-white/80 backdrop-blur-md sticky top-0 z-50">
        <div className="max-w-4xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-orange-600 rounded-xl flex items-center justify-center shadow-lg shadow-orange-200">
              <Eraser className="text-white w-6 h-6" />
            </div>
            <h1 className="text-xl font-semibold tracking-tight">PPT 水印消除大师</h1>
          </div>
          <div className="flex items-center gap-2 text-slate-400">
            <ShieldCheck className="w-4 h-4" />
            <span className="text-xs font-medium uppercase tracking-widest">Secure Local Processing</span>
          </div>
        </div>
      </nav>

      <main className="max-w-4xl mx-auto px-6 py-12">
        <div className="space-y-12">
          {/* Hero Section */}
          <section className="text-center space-y-4">
            <h2 className="text-4xl md:text-5xl font-bold tracking-tight text-slate-900">
              瞬间告别 <span className="text-orange-600 italic">烦人水印</span>
            </h2>
            <p className="text-slate-500 text-lg max-w-xl mx-auto">
              上传 PPT 文件，输入要删除的内容，我们将在全文档、母版和布局中为您深度清理。
            </p>
          </section>

          {/* Stepper / Main Interface */}
          <div className="bg-white border border-slate-100 rounded-3xl shadow-2xl shadow-slate-200/50 p-8 md:p-12">
            {!fileState ? (
              <motion.div
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className="space-y-8"
              >
                <div
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={onDrop}
                  onClick={() => fileInputRef.current?.click()}
                  className="group relative border-2 border-dashed border-slate-200 rounded-2xl p-12 transition-all hover:border-orange-400 hover:bg-orange-50/30 cursor-pointer text-center"
                >
                  <input
                    type="file"
                    ref={fileInputRef}
                    onChange={handleFileChange}
                    accept=".pptx"
                    className="hidden"
                  />
                  <div className="space-y-4">
                    <div className="w-16 h-16 bg-slate-50 group-hover:bg-orange-100 rounded-full flex items-center justify-center mx-auto transition-colors">
                      <Upload className="w-8 h-8 text-slate-400 group-hover:text-orange-600 transition-colors" />
                    </div>
                    <div className="space-y-1">
                      <p className="text-lg font-medium">点击或将文件拖拽到此处</p>
                      <p className="text-sm text-slate-400">支持 .pptx 格式 (PowerPoint 2007+)</p>
                    </div>
                  </div>
                </div>
              </motion.div>
            ) : (
              <motion.div
                initial={{ opacity: 0, scale: 0.98 }}
                animate={{ opacity: 1, scale: 1 }}
                className="space-y-8"
              >
                {/* File Preview */}
                <div className="flex items-center justify-between p-4 bg-slate-50 rounded-2xl border border-slate-100">
                  <div className="flex items-center gap-4">
                    <div className="w-12 h-12 bg-orange-100 rounded-lg flex items-center justify-center">
                      <FileDown className="text-orange-600 w-6 h-6" />
                    </div>
                    <div>
                      <h3 className="font-semibold text-slate-900 truncate max-w-[200px] md:max-w-md">{fileState.name}</h3>
                      <p className="text-xs text-slate-400">{fileState.size}</p>
                    </div>
                  </div>
                  {!isProcessing && !isDone && (
                    <button
                      onClick={reset}
                      className="p-2 hover:bg-white hover:text-red-500 rounded-lg transition-all text-slate-300"
                    >
                      <Trash2 className="w-5 h-5" />
                    </button>
                  )}
                </div>

                {!isDone ? (
                  <div className="space-y-6">
                    <div className="space-y-2">
                      <label className="text-sm font-semibold text-slate-700 ml-1">水印文字内容</label>
                      <div className="relative">
                        <input
                          type="text"
                          value={watermark}
                          onChange={(e) => setWatermark(e.target.value)}
                          placeholder="例如: abcccc 或 www.example.com"
                          className="w-full px-6 py-4 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-orange-100 focus:border-orange-500 transition-all outline-none text-lg"
                        />
                        <div className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-300">
                          <Eraser className="w-5 h-5" />
                        </div>
                      </div>
                      <p className="text-xs text-slate-400 ml-2">提示：请输入在 PPT 页面右下角或固定位置出现的原文文字。</p>
                    </div>

                    <button
                      onClick={removeWatermark}
                      disabled={!watermark || isProcessing}
                      className={`w-full py-4 rounded-2xl font-bold text-lg flex items-center justify-center gap-3 transition-all transform active:scale-[0.98] ${
                        !watermark || isProcessing
                          ? 'bg-slate-100 text-slate-400 cursor-not-allowed'
                          : 'bg-orange-600 text-white shadow-xl shadow-orange-200 hover:bg-orange-700'
                      }`}
                    >
                      {isProcessing ? (
                        <>
                          <Loader2 className="w-6 h-6 animate-spin" />
                          正在扫描并消除...
                        </>
                      ) : (
                        <>
                          <Eraser className="w-6 h-6" />
                          开始清理
                        </>
                      )}
                    </button>
                  </div>
                ) : (
                  <div className="space-y-6">
                    <div className="p-8 bg-green-50 border border-green-100 rounded-2xl text-center space-y-4">
                      <div className="w-16 h-16 bg-green-500 rounded-full flex items-center justify-center mx-auto shadow-lg shadow-green-100">
                        <CheckCircle2 className="text-white w-10 h-10" />
                      </div>
                      <div className="space-y-1">
                        <h3 className="text-2xl font-bold text-green-900">清理完成！</h3>
                        <p className="text-green-700">
                          已为您成功从 <span className="font-mono font-bold bg-green-200 px-2 py-0.5 rounded">{removedCount}</span> 处位置消除了水印内容。
                        </p>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <button
                        onClick={downloadFile}
                        className="w-full py-4 bg-orange-600 text-white rounded-2xl font-bold text-lg flex items-center justify-center gap-3 shadow-xl shadow-orange-200 hover:bg-orange-700 transition-all active:scale-[0.98]"
                      >
                        <FileDown className="w-6 h-6" />
                        下载修改后的文件
                      </button>
                      <button
                        onClick={reset}
                        className="w-full py-4 bg-white border border-slate-200 text-slate-600 rounded-2xl font-bold text-lg flex items-center justify-center gap-3 hover:bg-slate-50 transition-all active:scale-[0.98]"
                      >
                        <Upload className="w-6 h-6" />
                        处理另一个文件
                      </button>
                    </div>
                  </div>
                )}
              </motion.div>
            )}

            {/* Error Message */}
            <AnimatePresence>
              {error && (
                <motion.div
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="mt-6 p-4 bg-red-50 border border-red-100 rounded-xl flex items-center gap-3 text-red-600"
                >
                  <AlertCircle className="w-5 h-5 shrink-0" />
                  <p className="text-sm font-medium">{error}</p>
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          <section className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {[
              { icon: ShieldCheck, title: "隐私安全", desc: "所有处理均在您的浏览器本地完成，文件无需上传到第三方服务器。" },
              { icon: Eraser, title: "深度清理", desc: "不仅是幻灯片正文，母版和幻灯片布局中的隐形水印也能干净去除。" },
              { icon: CheckCircle2, title: "完美兼容", desc: "导出后仍保持原始格式与排版，无需担心样式错位问题。" }
            ].map((feature, i) => (
              <div key={i} className="p-6 bg-slate-50 rounded-2xl border border-slate-100 space-y-3">
                <div className="w-10 h-10 bg-white rounded-xl flex items-center justify-center shadow-sm">
                  <feature.icon className="w-5 h-5 text-orange-600" />
                </div>
                <h4 className="font-bold text-slate-900">{feature.title}</h4>
                <p className="text-sm text-slate-500 leading-relaxed">{feature.desc}</p>
              </div>
            ))}
          </section>
        </div>
      </main>

      <footer className="max-w-4xl mx-auto px-6 py-12 border-t border-slate-100 text-center space-y-2">
        <p className="text-slate-400 text-sm italic">
          注：本工具仅用于合法合规的文档编辑用途，请确保您拥有原文件的相应操作权利。
        </p>
        <p className="text-slate-300 text-xs">
          Built with React & JSZip &middot; 无中间层数据传输
        </p>
      </footer>
    </div>
  );
}
