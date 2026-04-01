import React, { useState, useRef, useEffect } from 'react';
import {
  Upload,
  FileText,
  Loader2,
  Trash2,
  ChevronDown,
  ChevronUp,
  AlertTriangle,
  CheckCircle2,
  FileSpreadsheet,
  Zap,
  Info,
  Download,
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { analyzeSourceFiles, ConsolidationResult } from './lib/excelService';
import {
  consolidateToMatrix,
  ConsolidationReport,
  MatchReport,
  MatchedArticle,
} from './lib/consolidateService';

type TabType = 'analyze' | 'consolidate';

interface SourceFileStatus {
  file: File;
  completed: boolean;
}

export default function App() {
  // Shared
  const [activeTab, setActiveTab] = useState<TabType>('consolidate');
  const [files, setFiles] = useState<SourceFileStatus[]>([]);
  const sourceInputRef = useRef<HTMLInputElement>(null);

  // Analyze tab
  const [isProcessing, setIsProcessing] = useState(false);
  const [results, setResults] = useState<ConsolidationResult[] | null>(null);
  const [expandedIdx, setExpandedIdx] = useState<number | null>(null);

  // Consolidate tab
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [templateLoading, setTemplateLoading] = useState(true);

  // Charger le template automatiquement au démarrage
  useEffect(() => {
    fetch('/matrice-v3.1.xlsx')
      .then((res) => res.blob())
      .then((blob) => {
        const file = new File([blob], 'Fichier Global 2026 MATRICE V3.1 YM.xlsx', {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        setTemplateFile(file);
      })
      .catch(() => {
        // Si le fetch échoue, l'utilisateur peut uploader manuellement
      })
      .finally(() => setTemplateLoading(false));
  }, []);
  const [isConsolidating, setIsConsolidating] = useState(false);
  const [consolidationProgress, setConsolidationProgress] = useState<{
    current: number;
    total: number;
    clientName: string;
  } | null>(null);
  const [consolidationReport, setConsolidationReport] =
    useState<ConsolidationReport | null>(null);
  const [consolidationBlob, setConsolidationBlob] = useState<Blob | null>(null);
  const [expandedReportIdx, setExpandedReportIdx] = useState<number | null>(
    null
  );

  // ── Handlers ────────────────────────────────────────────────
  const handleSourceUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files).map((file) => ({
        file,
        completed: false,
      }));
      setFiles((prev) => [...prev, ...newFiles]);
    }
  };

  const removeSourceFile = (index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
    if (results) {
      setResults((prev) => (prev ? prev.filter((_, i) => i !== index) : null));
    }
    if (sourceInputRef.current) sourceInputRef.current.value = '';
  };

  const toggleCompleted = (index: number) => {
    setFiles((prev) =>
      prev.map((f, i) =>
        i === index ? { ...f, completed: !f.completed } : f
      )
    );
  };

  // Analyze
  const startAnalysis = async () => {
    if (files.length === 0) return;
    setIsProcessing(true);
    setResults(null);
    try {
      const res = await analyzeSourceFiles(files.map((f) => f.file));
      setResults(res);
    } catch (error) {
      console.error(error);
      alert("Une erreur est survenue lors de l'analyse.");
    } finally {
      setIsProcessing(false);
    }
  };

  // Consolidate
  const startConsolidation = async () => {
    if (files.length === 0 || !templateFile) return;
    setIsConsolidating(true);
    setConsolidationReport(null);
    setConsolidationProgress(null);

    try {
      const { blob, report } = await consolidateToMatrix(
        files.map((f) => f.file),
        templateFile,
        (current, total, clientName) => {
          setConsolidationProgress({ current, total, clientName });
        }
      );

      setConsolidationReport(report);
      setConsolidationBlob(blob);
    } catch (error) {
      console.error(error);
      alert('Erreur lors de la consolidation : ' + (error as Error).message);
    } finally {
      setIsConsolidating(false);
      setConsolidationProgress(null);
    }
  };

  const toggleExpand = (idx: number) =>
    setExpandedIdx(expandedIdx === idx ? null : idx);
  const toggleReportExpand = (idx: number) =>
    setExpandedReportIdx(expandedReportIdx === idx ? null : idx);

  const colorBadge = (color: string) => {
    const c = color.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (c.includes('red') || c.includes('rouge'))
      return 'bg-red-100 text-red-700 border border-red-200';
    if (c.includes('rose') || c.includes('pink'))
      return 'bg-pink-100 text-pink-700 border border-pink-200';
    if (c.includes('white') || c.includes('blanc'))
      return 'bg-amber-50 text-amber-700 border border-amber-200';
    if (c.includes('champagne') || c.includes('sparkling') || c.includes('bulles'))
      return 'bg-yellow-50 text-yellow-700 border border-yellow-200';
    return 'bg-slate-100 text-slate-600 border border-slate-200';
  };

  const progressPct = consolidationProgress
    ? Math.round(
        (consolidationProgress.current / consolidationProgress.total) * 100
      )
    : 0;

  // ── Render ──────────────────────────────────────────────────
  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans selection:bg-brand/20 selection:text-brand">
      <div className="max-w-7xl mx-auto px-4 py-8">
        {/* ── Header ────────────────────────────────── */}
        <header className="mb-6 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-brand/10 flex items-center justify-center">
              <FileSpreadsheet className="text-brand" size={20} />
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight">
                Peuch &amp; Besse
              </h1>
              <p className="text-xs text-slate-400">
                Analyse &amp; Consolidation Excel Ambassade
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => {
                setFiles([]);
                setResults(null);
                setConsolidationReport(null);
                setConsolidationBlob(null);
                if (sourceInputRef.current) sourceInputRef.current.value = '';
              }}
              className="px-4 py-2 text-xs font-semibold text-slate-500 hover:bg-slate-100 rounded-lg transition-colors border border-slate-200"
            >
              Réinitialiser
            </button>
          </div>
        </header>

        {/* ── Tabs ──────────────────────────────────── */}
        <div className="flex gap-1 mb-6 bg-white p-1 rounded-lg border border-slate-200 shadow-sm w-fit">
          <button
            onClick={() => setActiveTab('consolidate')}
            className={`px-5 py-2.5 text-xs font-bold rounded-md transition-all flex items-center gap-2 ${
              activeTab === 'consolidate'
                ? 'bg-brand text-white shadow-sm'
                : 'text-slate-500 hover:bg-slate-50'
            }`}
          >
            <Zap size={14} />
            Consolidation Matrice
          </button>
          <button
            onClick={() => setActiveTab('analyze')}
            className={`px-5 py-2.5 text-xs font-bold rounded-md transition-all flex items-center gap-2 ${
              activeTab === 'analyze'
                ? 'bg-brand text-white shadow-sm'
                : 'text-slate-500 hover:bg-slate-50'
            }`}
          >
            <FileText size={14} />
            Analyse Individuelle
          </button>
        </div>

        {/* ── Main Grid ─────────────────────────────── */}
        <main className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          {/* ── Left: Source Files ────────────────────── */}
          <section className="lg:col-span-4 space-y-4">
            {/* Source files upload */}
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="px-5 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500">
                  Fichiers Sources
                </h2>
                <span className="text-xs font-bold text-brand px-2 py-0.5 bg-brand/10 rounded-full">
                  {files.length} fichier{files.length > 1 ? 's' : ''}
                </span>
              </div>

              <div className="p-4 space-y-3">
                <div
                  onClick={() => sourceInputRef.current?.click()}
                  className="border-2 border-dashed border-slate-200 rounded-lg p-6 text-center hover:border-brand/50 hover:bg-brand/5 transition-all cursor-pointer group"
                >
                  <Upload
                    className="mx-auto mb-2 text-slate-400 group-hover:text-brand transition-colors"
                    size={24}
                  />
                  <p className="text-xs font-semibold text-slate-500 group-hover:text-brand">
                    Ajouter des fichiers Excel
                  </p>
                  <p className="text-[10px] text-slate-400 mt-1">
                    .xlsx, .xls, .xlsm, .xlsb
                  </p>
                  <input
                    type="file"
                    multiple
                    accept=".xlsx,.xlsm,.xlsb,.xls"
                    className="hidden"
                    ref={sourceInputRef}
                    onChange={handleSourceUpload}
                  />
                </div>

                <div className="space-y-2 max-h-[400px] overflow-y-auto custom-scrollbar">
                  <AnimatePresence initial={false}>
                    {files.map((f, idx) => (
                      <motion.div
                        key={f.file.name + idx}
                        layout
                        initial={{ opacity: 0, y: -4 }}
                        animate={{ opacity: 1, y: 0 }}
                        className={`flex items-center gap-3 p-3 rounded-lg border transition-all ${
                          f.completed
                            ? 'bg-slate-50 border-slate-100 opacity-50'
                            : 'bg-white border-slate-200 hover:border-brand/30'
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={f.completed}
                          onChange={() => toggleCompleted(idx)}
                          className="w-4 h-4 rounded border-slate-300 text-brand focus:ring-brand accent-brand cursor-pointer"
                        />
                        <div className="flex-1 min-w-0">
                          <p
                            className={`text-sm font-medium truncate ${
                              f.completed
                                ? 'text-slate-400 line-through'
                                : 'text-slate-900'
                            }`}
                          >
                            {f.file.name}
                          </p>
                          <p className="text-[10px] font-mono text-slate-400 uppercase">
                            {(f.file.size / 1024).toFixed(0)} KB
                          </p>
                        </div>
                        <button
                          onClick={() => removeSourceFile(idx)}
                          className="text-slate-300 hover:text-rose-500 transition-colors"
                        >
                          <Trash2 size={14} />
                        </button>
                      </motion.div>
                    ))}
                  </AnimatePresence>
                </div>
              </div>
            </div>


            {/* Action Button */}
            <div>
              {activeTab === 'consolidate' ? (
                <button
                  disabled={
                    isConsolidating || files.length === 0 || !templateFile
                  }
                  onClick={startConsolidation}
                  className="w-full px-6 py-3.5 bg-brand text-white text-sm font-bold rounded-xl hover:bg-brand/90 transition-all disabled:opacity-40 disabled:cursor-not-allowed flex items-center justify-center gap-2 shadow-md shadow-brand/20"
                >
                  {isConsolidating ? (
                    <Loader2 className="animate-spin" size={16} />
                  ) : (
                    <Zap size={16} />
                  )}
                  {isConsolidating
                    ? 'Consolidation en cours...'
                    : 'Générer la Matrice Consolidée'}
                </button>
              ) : (
                <button
                  disabled={isProcessing || files.length === 0}
                  onClick={startAnalysis}
                  className="w-full px-6 py-3.5 bg-brand text-white text-sm font-bold rounded-xl hover:bg-brand/90 transition-all disabled:opacity-40 disabled:cursor-not-allowed flex items-center justify-center gap-2 shadow-md shadow-brand/20"
                >
                  {isProcessing ? (
                    <Loader2 className="animate-spin" size={16} />
                  ) : (
                    <FileText size={16} />
                  )}
                  {isProcessing ? 'Analyse...' : 'Lancer l\'analyse'}
                </button>
              )}
            </div>

            {/* Progress bar */}
            {isConsolidating && consolidationProgress && (
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="bg-white rounded-xl border border-slate-200 shadow-sm p-4"
              >
                <div className="flex justify-between items-center mb-2">
                  <span className="text-[10px] font-bold uppercase text-slate-400">
                    Progression
                  </span>
                  <span className="text-xs font-bold text-brand">
                    {progressPct}%
                  </span>
                </div>
                <div className="h-2 bg-slate-100 rounded-full overflow-hidden">
                  <motion.div
                    className="h-full bg-brand rounded-full"
                    initial={{ width: 0 }}
                    animate={{ width: `${progressPct}%` }}
                    transition={{ ease: 'easeOut' }}
                  />
                </div>
                <p className="text-[10px] text-slate-400 mt-2 truncate">
                  {consolidationProgress.clientName}
                </p>
              </motion.div>
            )}
          </section>

          {/* ── Right: Results ───────────────────────── */}
          <section className="lg:col-span-8 space-y-4">
            {activeTab === 'consolidate' ? (
              /* ── Consolidation Report ─────────────── */
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="px-5 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                  <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500">
                    Rapport de Consolidation
                  </h2>
                  {consolidationReport && (
                    <div className="flex items-center gap-3">
                      <span className="text-xs font-bold text-brand">
                        {consolidationReport.totalClients} clients
                      </span>
                      {consolidationBlob && (
                        <button
                          onClick={() => {
                            const url = URL.createObjectURL(consolidationBlob);
                            const a = document.createElement('a');
                            a.href = url;
                            a.download = 'GLOBAL_CONSOLIDE.xlsx';
                            document.body.appendChild(a);
                            a.click();
                            document.body.removeChild(a);
                            URL.revokeObjectURL(url);
                          }}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-brand text-white text-xs font-bold rounded-lg hover:bg-brand/90 transition-colors shadow-sm"
                        >
                          <Download size={13} />
                          Télécharger
                        </button>
                      )}
                    </div>
                  )}
                </div>

                <div className="p-4">
                  {!consolidationReport ? (
                    <div className="py-20 text-center">
                      <Zap
                        className="mx-auto mb-3 text-slate-200"
                        size={40}
                      />
                      <p className="text-sm font-medium text-slate-400">
                        Ajoutez vos fichiers sources et lancez la consolidation
                      </p>
                      <div className="mt-6 max-w-md mx-auto text-left space-y-3">
                        <div className="flex items-start gap-3 p-3 rounded-lg bg-slate-50">
                          <span className="w-6 h-6 rounded-full bg-brand/10 text-brand text-xs font-bold flex items-center justify-center shrink-0">
                            1
                          </span>
                          <p className="text-xs text-slate-500">
                            Ajoutez vos <strong>fichiers sources</strong>{' '}
                            (commandes individuelles par client)
                          </p>
                        </div>
                        <div className="flex items-start gap-3 p-3 rounded-lg bg-slate-50">
                          <span className="w-6 h-6 rounded-full bg-brand/10 text-brand text-xs font-bold flex items-center justify-center shrink-0">
                            2
                          </span>
                          <p className="text-xs text-slate-500">
                            Cliquez{' '}
                            <strong>"Générer la Matrice Consolidée"</strong> →
                            téléchargement automatique
                          </p>
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-4">
                      {/* Summary cards */}
                      <div className="grid grid-cols-3 gap-3">
                        <div className="p-4 rounded-lg bg-emerald-50 border border-emerald-100 text-center">
                          <p className="text-2xl font-bold text-emerald-700">
                            {consolidationReport.totalMatched}
                          </p>
                          <p className="text-[10px] font-bold uppercase text-emerald-500 mt-1">
                            Produits matchés
                          </p>
                        </div>
                        <div className="p-4 rounded-lg bg-amber-50 border border-amber-100 text-center">
                          <p className="text-2xl font-bold text-amber-700">
                            {consolidationReport.totalUnmatched}
                          </p>
                          <p className="text-[10px] font-bold uppercase text-amber-500 mt-1">
                            Non trouvés
                          </p>
                        </div>
                        <div className="p-4 rounded-lg bg-brand/5 border border-brand/20 text-center">
                          <p className="text-2xl font-bold text-brand">
                            {consolidationReport.totalClients}
                          </p>
                          <p className="text-[10px] font-bold uppercase text-brand/70 mt-1">
                            Clients traités
                          </p>
                        </div>
                      </div>

                      {/* Per-client reports */}
                      <div className="space-y-2">
                        {consolidationReport.reports.map(
                          (rep: MatchReport, idx: number) => (
                            <div
                              key={idx}
                              className="rounded-lg border border-slate-200 bg-white hover:border-brand/20 transition-all shadow-sm"
                            >
                              <div
                                className="p-4 flex items-center justify-between gap-4 cursor-pointer"
                                onClick={() => toggleReportExpand(idx)}
                              >
                                <div className="flex items-center gap-3 flex-1 min-w-0">
                                  {rep.unmatched.length === 0 ? (
                                    <CheckCircle2
                                      className="text-emerald-500 shrink-0"
                                      size={18}
                                    />
                                  ) : (
                                    <AlertTriangle
                                      className="text-amber-500 shrink-0"
                                      size={18}
                                    />
                                  )}
                                  <div className="min-w-0">
                                    <p className="text-sm font-bold text-slate-900 truncate">
                                      {rep.clientName}
                                    </p>
                                    <p className="text-[10px] font-mono text-slate-400 truncate uppercase">
                                      {rep.fileName}
                                    </p>
                                  </div>
                                </div>

                                <div className="flex items-center gap-4 shrink-0">
                                  <div className="text-right">
                                    <p className="text-sm font-bold text-slate-900">
                                      {rep.totalBottles} btl
                                    </p>
                                    <p className="text-[10px] text-slate-400">
                                      <span className="text-emerald-600 font-bold">
                                        {rep.matched}
                                      </span>{' '}
                                      matchés
                                      {rep.unmatched.length > 0 && (
                                        <>
                                          {' · '}
                                          <span className="text-amber-600 font-bold">
                                            {rep.unmatched.length}
                                          </span>{' '}
                                          manquants
                                        </>
                                      )}
                                    </p>
                                  </div>
                                  <button
                                    className={`p-1.5 rounded-lg transition-colors ${
                                      expandedReportIdx === idx
                                        ? 'bg-brand text-white'
                                        : 'bg-slate-100 text-slate-400'
                                    }`}
                                  >
                                    {expandedReportIdx === idx ? (
                                      <ChevronUp size={14} />
                                    ) : (
                                      <ChevronDown size={14} />
                                    )}
                                  </button>
                                </div>
                              </div>

                              <AnimatePresence>
                                {expandedReportIdx === idx && (
                                  <motion.div
                                    initial={{ height: 0, opacity: 0 }}
                                    animate={{ height: 'auto', opacity: 1 }}
                                    exit={{ height: 0, opacity: 0 }}
                                    className="border-t border-slate-100 overflow-hidden"
                                  >
                                    <div className="p-4 space-y-4">
                                      {/* Matched articles */}
                                      {rep.matchedArticles.length > 0 && (
                                        <div>
                                          <div className="flex items-center gap-2 mb-2">
                                            <CheckCircle2 className="text-emerald-500" size={13} />
                                            <p className="text-[10px] font-bold uppercase text-emerald-600">
                                              Produits matchés ({rep.matchedArticles.length})
                                            </p>
                                          </div>
                                          <table className="w-full text-left border-collapse">
                                            <thead>
                                              <tr className="text-[10px] uppercase font-bold text-slate-400 border-b border-slate-200">
                                                <th className="py-1.5 px-2">Désignation</th>
                                                <th className="py-1.5 px-2">Couleur</th>
                                                <th className="py-1.5 px-2">Millésime</th>
                                                <th className="py-1.5 px-2">Taille</th>
                                                <th className="py-1.5 px-2 text-right">Qté (btl)</th>
                                              </tr>
                                            </thead>
                                            <tbody>
                                              {rep.matchedArticles.map((art: MatchedArticle, aIdx: number) => (
                                                <tr key={aIdx} className="border-b border-slate-100 hover:bg-slate-50">
                                                  <td className="py-2 px-2">
                                                    <p className="text-xs font-semibold text-slate-900">{art.designation}</p>
                                                    {art.appellation && art.appellation !== art.designation && (
                                                      <p className="text-[10px] text-brand italic">{art.appellation}</p>
                                                    )}
                                                  </td>
                                                  <td className="py-2 px-2">
                                                    {art.color ? (
                                                      <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${colorBadge(art.color)}`}>
                                                        {art.color}
                                                      </span>
                                                    ) : (
                                                      <span className="text-slate-300">—</span>
                                                    )}
                                                  </td>
                                                  <td className="py-2 px-2">
                                                    {art.millesime ? (
                                                      <span className="text-xs font-mono font-semibold text-slate-700">{art.millesime}</span>
                                                    ) : (
                                                      <span className="text-slate-300">—</span>
                                                    )}
                                                  </td>
                                                  <td className="py-2 px-2">
                                                    {art.taille ? (
                                                      <span className="text-xs font-mono font-semibold text-slate-700">{art.taille}</span>
                                                    ) : (
                                                      <span className="text-slate-300">—</span>
                                                    )}
                                                  </td>
                                                  <td className="py-2 px-2 text-right font-mono text-sm font-bold text-slate-900">
                                                    {art.qty}
                                                  </td>
                                                </tr>
                                              ))}
                                            </tbody>
                                          </table>
                                        </div>
                                      )}

                                      {/* Unmatched articles */}
                                      {rep.unmatched.length > 0 && (
                                        <div>
                                          <div className="flex items-center gap-2 mb-2">
                                            <Info className="text-amber-500" size={13} />
                                            <p className="text-[10px] font-bold uppercase text-amber-600">
                                              Non trouvés dans la Matrice ({rep.unmatched.length})
                                            </p>
                                          </div>
                                          <div className="space-y-1">
                                            {rep.unmatched.map((name: string, uIdx: number) => (
                                              <p key={uIdx} className="text-xs text-amber-800 font-medium py-1 px-2 bg-amber-100/50 rounded">
                                                {name}
                                              </p>
                                            ))}
                                          </div>
                                        </div>
                                      )}
                                    </div>
                                  </motion.div>
                                )}
                              </AnimatePresence>
                            </div>
                          )
                        )}
                      </div>
                    </div>
                  )}
                </div>
              </div>
            ) : (
              /* ── Analysis Results (existing) ─────────── */
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="px-5 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                  <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500">
                    Résultats d'Analyse
                  </h2>
                  {results && (
                    <span className="text-xs font-bold text-brand">
                      {results.length} rapports
                    </span>
                  )}
                </div>

                <div className="p-4">
                  {!results ? (
                    <div className="py-20 text-center">
                      <FileText
                        className="mx-auto mb-3 text-slate-200"
                        size={40}
                      />
                      <p className="text-sm font-medium text-slate-400">
                        Lancez l'analyse pour voir les résultats
                      </p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {results.map((res, idx) => (
                        <div
                          key={idx}
                          className={`rounded-lg border transition-all ${
                            files[idx]?.completed
                              ? 'bg-slate-50 border-slate-100 opacity-60'
                              : 'bg-white border-slate-200 hover:border-brand/20 shadow-sm'
                          }`}
                        >
                          <div className="p-4 flex items-center justify-between gap-4">
                            <div className="flex items-center gap-4 flex-1 min-w-0">
                              <input
                                type="checkbox"
                                checked={files[idx]?.completed || false}
                                onChange={() => toggleCompleted(idx)}
                                className="w-4 h-4 rounded border-slate-300 text-brand focus:ring-brand accent-brand cursor-pointer"
                              />
                              <div className="min-w-0">
                                <h3
                                  className={`text-lg font-bold truncate ${
                                    files[idx]?.completed
                                      ? 'text-slate-400'
                                      : 'text-slate-900'
                                  }`}
                                >
                                  {res.clientName}
                                </h3>
                                <p className="text-xs text-slate-400 truncate font-mono uppercase tracking-tight">
                                  {res.fileName}
                                </p>
                              </div>
                            </div>

                            <div className="flex items-center gap-6 shrink-0">
                              <div className="text-right">
                                <p className="text-xl font-bold text-slate-900">
                                  {res.globalTotal.toLocaleString('fr-FR', {
                                    style: 'currency',
                                    currency: 'EUR',
                                  })}
                                </p>
                                <p className="text-[10px] font-bold text-brand uppercase">
                                  {res.articles?.length || 0} articles
                                </p>
                              </div>
                              <button
                                onClick={() => toggleExpand(idx)}
                                className={`p-2 rounded-lg transition-colors ${
                                  expandedIdx === idx
                                    ? 'bg-brand text-white'
                                    : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                                }`}
                              >
                                {expandedIdx === idx ? (
                                  <ChevronUp size={18} />
                                ) : (
                                  <ChevronDown size={18} />
                                )}
                              </button>
                            </div>
                          </div>

                          <AnimatePresence>
                            {expandedIdx === idx && (
                              <motion.div
                                initial={{ height: 0, opacity: 0 }}
                                animate={{ height: 'auto', opacity: 1 }}
                                exit={{ height: 0, opacity: 0 }}
                                className="border-t border-slate-100 bg-slate-50/50 overflow-hidden"
                              >
                                <div className="p-4">
                                  {res.error ? (
                                    <div className="p-4 bg-rose-50 border border-rose-100 rounded-lg flex items-start gap-3">
                                      <AlertTriangle
                                        className="text-rose-500 shrink-0"
                                        size={18}
                                      />
                                      <p className="text-sm text-rose-700 font-medium">
                                        {res.error}
                                      </p>
                                    </div>
                                  ) : (
                                    <div className="space-y-4">
                                      <div className="overflow-x-auto">
                                        <table className="w-full text-left border-collapse">
                                          <thead>
                                            <tr className="text-[10px] uppercase font-bold text-slate-400 border-b border-slate-200">
                                              <th className="py-2 px-2">
                                                Désignation
                                              </th>
                                              <th className="py-2 px-2">
                                                Couleur
                                              </th>
                                              <th className="py-2 px-2">
                                                Millésime
                                              </th>
                                              <th className="py-2 px-2">
                                                Taille
                                              </th>
                                              <th className="py-2 px-2 text-right">
                                                Qté
                                              </th>
                                              <th className="py-2 px-2 text-right">
                                                Total HT
                                              </th>
                                            </tr>
                                          </thead>
                                          <tbody className="text-sm">
                                            {res.articles?.map((art, aIdx) => (
                                              <tr
                                                key={aIdx}
                                                className="border-b border-slate-100 hover:bg-white transition-colors"
                                              >
                                                <td className="py-3 px-2">
                                                  <p className="font-bold text-slate-900">
                                                    {art.designation}
                                                  </p>
                                                  <p className="text-[10px] text-brand font-medium italic">
                                                    {art.appellation}
                                                  </p>
                                                </td>
                                                <td className="py-3 px-2">
                                                  {art.color ? (
                                                    <span className={`text-xs font-semibold px-2.5 py-1 rounded-full ${colorBadge(art.color)}`}>
                                                      {art.color}
                                                    </span>
                                                  ) : (
                                                    <span className="text-slate-300">—</span>
                                                  )}
                                                </td>
                                                <td className="py-3 px-2">
                                                  {art.millesime ? (
                                                    <span className="text-xs font-mono font-semibold text-slate-700">{art.millesime}</span>
                                                  ) : (
                                                    <span className="text-slate-300">—</span>
                                                  )}
                                                </td>
                                                <td className="py-3 px-2">
                                                  {art.taille ? (
                                                    <span className="text-xs font-mono font-semibold text-slate-700">{art.taille}</span>
                                                  ) : (
                                                    <span className="text-slate-300">—</span>
                                                  )}
                                                </td>
                                                <td className="py-3 px-2 text-right font-mono text-slate-500">
                                                  {art.quantity}
                                                </td>
                                                <td className="py-3 px-2 text-right font-bold text-slate-900">
                                                  {art.total.toLocaleString(
                                                    'fr-FR',
                                                    {
                                                      style: 'currency',
                                                      currency: 'EUR',
                                                    }
                                                  )}
                                                </td>
                                              </tr>
                                            ))}
                                          </tbody>
                                        </table>
                                      </div>

                                      <div className="flex flex-col items-end gap-2 pt-4 border-t border-slate-200">
                                        <div className="flex justify-between w-full max-w-[200px] text-xs text-slate-500">
                                          <span>Bouteilles</span>
                                          <span className="font-bold text-slate-900">
                                            {res.totalBottles}
                                          </span>
                                        </div>
                                        <div className="flex justify-between w-full max-w-[200px] text-xs text-slate-500">
                                          <span>Transport</span>
                                          <span className="font-bold text-slate-900">
                                            {res.transport.toLocaleString(
                                              'fr-FR',
                                              {
                                                style: 'currency',
                                                currency: 'EUR',
                                              }
                                            )}
                                          </span>
                                        </div>
                                        <div className="flex justify-between w-full max-w-[200px] items-center pt-2 mt-1 border-t-2 border-brand">
                                          <span className="text-xs font-bold text-brand uppercase">
                                            Total Net
                                          </span>
                                          <span className="text-xl font-bold text-slate-900">
                                            {res.globalTotal.toLocaleString(
                                              'fr-FR',
                                              {
                                                style: 'currency',
                                                currency: 'EUR',
                                              }
                                            )}
                                          </span>
                                        </div>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              </motion.div>
                            )}
                          </AnimatePresence>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}
          </section>
        </main>
      </div>

    </div>
  );
}
