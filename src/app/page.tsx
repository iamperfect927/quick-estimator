'use client';

import { useChat } from '@ai-sdk/react';
import { DefaultChatTransport, UIMessage } from 'ai';
import { useState, ChangeEvent, FormEvent } from 'react';

export default function EstimatorDashboard() {
  const [calculationResult, setCalculationResult] = useState<any>(null);
  
  // File upload states
  const [fieldStudyFile, setFieldStudyFile] = useState<File | null>(null);
  const [priceListFile, setPriceListFile] = useState<File | null>(null);
  const [uploadStatus, setUploadStatus] = useState<string>('');
  
  // Local input text state (Required by AI SDK v6 since input/handleInputChange/handleSubmit are separate)
  const [inputText, setInputText] = useState('');

  // Google Drive Modal & Fetch states
  const [driveModalOpen, setDriveModalOpen] = useState(false);
  const [driveImportType, setDriveImportType] = useState<'field' | 'price' | null>(null);
  const [driveUrlInput, setDriveUrlInput] = useState('');
  const [driveLoading, setDriveLoading] = useState(false);
  const [driveError, setDriveError] = useState('');
  const [helpModalOpen, setHelpModalOpen] = useState(false);

  // Initializing useChat with DefaultChatTransport for custom endpoint and onToolCall handler
  const { messages, sendMessage, setMessages, status } = useChat({
    transport: new DefaultChatTransport({ api: '/api/chat' }),
    onToolCall({ toolCall }) {
      if (toolCall.toolName === 'calculateEstimate') {
        const input = toolCall.input as any;
        setCalculationResult((prev: any) => ({
          ...input,
          customerName: input.customerName || prev?.customerName,
          metrics: input.metrics || prev?.metrics,
          apartments: input.apartments || prev?.apartments
        }));
      }
    }
  });

  // Handle file selection
  const handleFileChange = (e: ChangeEvent<HTMLInputElement>, type: 'field' | 'price') => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (type === 'field') setFieldStudyFile(file);
      if (type === 'price') setPriceListFile(file);
      setUploadStatus('');
    }
  };

  // Trigger Google Drive share link dialog modal
  const handleGoogleDrivePicker = (type: 'field' | 'price') => {
    setDriveImportType(type);
    setDriveUrlInput('');
    setDriveError('');
    setDriveModalOpen(true);
  };

  // Import file from Google Drive via backend proxy endpoint
  const handleImportFromDrive = async () => {
    if (!driveUrlInput.trim()) {
      setDriveError('⚠️ Please enter a Google Drive sharing link.');
      return;
    }

    setDriveLoading(true);
    setDriveError('');

    try {
      const res = await fetch(`/api/fetch-drive-file?url=${encodeURIComponent(driveUrlInput.trim())}`);
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || 'Failed to download file from Google Drive.');
      }

      const blob = await res.blob();
      
      const contentDisposition = res.headers.get('content-disposition');
      let fileName = driveImportType === 'field' ? 'field_study.xlsx' : 'price_list.xlsx';
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="?([^"]+)"?/);
        if (match && match[1]) fileName = match[1];
      }

      const file = new File([blob], fileName, { type: blob.type });

      if (driveImportType === 'field') {
        setFieldStudyFile(file);
      } else if (driveImportType === 'price') {
        setPriceListFile(file);
      }

      setDriveModalOpen(false);
      setUploadStatus(`✅ Imported "${fileName}" from Google Drive!`);
    } catch (err: any) {
      console.error('[Google Drive Import Error]:', err);
      setDriveError(err.message || 'An error occurred while fetching the file.');
    } finally {
      setDriveLoading(false);
    }
  };

  // Process files: send to backend parser then inject results into state + AI context
  const handleAttachFilesToAI = async () => {
    if (!fieldStudyFile || !priceListFile) {
      setUploadStatus('❌ Please upload both required files first.');
      return;
    }

    setUploadStatus('⏳ Parsing load profiles and cross-referencing catalog...');

    try {
      // Build FormData payload with both workbooks
      const formData = new FormData();
      formData.append('fieldStudy', fieldStudyFile);
      formData.append('priceList', priceListFile);

      const res = await fetch('/api/parse-solar-files', {
        method: 'POST',
        body: formData
      });

      if (!res.ok) {
        const err = await res.json();
        setUploadStatus(`❌ Parse Error: ${err.error || 'Unknown failure'}`);
        return;
      }

      const parsed = await res.json();

      // Feed aggregated BOQ directly into the right-panel calculation state
      setCalculationResult({
        materials: parsed.materials,
        labor: parsed.labor,
        marginPercentage: parsed.marginPercentage,
        customerName: parsed.customerName,
        metrics: parsed.metrics,
        apartments: parsed.apartments
      });

      // Format a detailed system notice so the AI agent is aware of all computed values
      const m = parsed.metrics;
      const warningsBlock = parsed.warnings?.length
        ? `\n⚠️ SYSTEM WARNINGS:\n${parsed.warnings.map((w: string) => `  • ${w}`).join('\n')}`
        : '';

      const totalMat = parsed.materials?.reduce((acc: number, item: any) => acc + item.quantity * item.unitPrice, 0) || 0;
      const totalLab = parsed.labor?.reduce((acc: number, item: any) => acc + item.hours * item.hourlyRate, 0) || 0;
      const grandTotalXAF = totalMat + totalLab;

      const systemNoticeText = `[SYSTEM NOTICE: Ingestion of Tollgate load profiles complete for "${parsed.customerName}".

📊 BUILDING AGGREGATION METRICS:
  • Apartments Parsed: ${m?.apartmentCount ?? '?'}
  • Total Devices: ${m?.totalDeviceCount ?? '?'}
  • Peak Load: ${m?.peakKW?.toFixed(2) ?? '?'} kW
  • Day Consumption: ${m?.dayConsumptionKWh?.toFixed(2) ?? '?'} kWh
  • Night Consumption: ${m?.nightConsumptionKWh?.toFixed(2) ?? '?'} kWh

📦 SELECTED BOQ (${parsed.materials?.length ?? 0} line items):
${parsed.materials?.map((item: any) => `  • ${item.name} — Qty: ${item.quantity} @ ${item.unitPrice.toLocaleString()} XAF`).join('\n') ?? ''}

🛠️ LABOR COST RULE: Labor = 30% of total equipment cost.
  • Equipment Subtotal: ${totalMat.toLocaleString()} XAF
  • Labor (30%): ${totalLab.toLocaleString()} XAF

💰 FINANCIAL SUMMARY:
  • Materials Subtotal: ${totalMat.toLocaleString()} XAF
  • Labor Subtotal (30% of Equipment): ${totalLab.toLocaleString()} XAF
  • GRAND TOTAL: ${grandTotalXAF.toLocaleString(undefined, { maximumFractionDigits: 0 })} XAF${warningsBlock}

The right-panel estimation dashboard has been auto-populated. You can now answer follow-up questions, adjust margins, or generate the Excel report.]`;

      const systemNoticeMsg: UIMessage = {
        id: 'solar-ingestion-' + Date.now(),
        role: 'user',
        parts: [{ type: 'text', text: systemNoticeText }]
      };

      setMessages((prev) => [...prev, systemNoticeMsg]);
      setUploadStatus(`✅ Parsed ${m?.apartmentCount ?? 0} apartments · ${m?.totalDeviceCount ?? 0} devices · ${m?.peakKW?.toFixed(1) ?? '?'} kW peak. Right panel updated!`);
    } catch (err: any) {
      console.error('File ingestion error:', err);
      setUploadStatus(`❌ Connection error: ${err.message}`);
    }
  };

  // Triggered when form is submitted
  const handleSubmit = (e: FormEvent) => {
    e.preventDefault();
    if (!inputText.trim()) return;
    sendMessage({ text: inputText });
    setInputText('');
  };

  // Triggered when a quick suggestion pill is clicked
  const handleSuggestionClick = (suggestionText: string) => {
    sendMessage({ text: suggestionText });
  };

  const downloadExcel = async () => {
    if (!calculationResult) return;
    
    try {
      const response = await fetch('/api/download-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(calculationResult),
      });

      if (!response.ok) throw new Error('Generation failed');

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Solar_Estimate_${Date.now()}.xlsx`;
      a.click();
    } catch (err) {
      console.error(err);
      alert('Failed to download Excel file. Please verify calculation parameters.');
    }
  };

  // Helper to extract text safely from UIMessage parts
  const getMessageTextContent = (m: UIMessage): string => {
    return m.parts
      .filter((p) => p.type === 'text')
      .map((p) => (p as any).text || '')
      .join('');
  };

  // Filter messages to hide system notice tokens for clean user viewing
  const visibleMessages = messages.filter((m) => {
    const text = getMessageTextContent(m);
    return !text.includes('[SYSTEM NOTICE');
  });

  // Calculate quick totals for visual preview
  const materialsSubtotal = calculationResult?.materials?.reduce((acc: number, item: any) => acc + (item.quantity * item.unitPrice), 0) || 0;
  const laborSubtotal = calculationResult?.labor?.reduce((acc: number, item: any) => acc + (item.hours * item.hourlyRate), 0) || 0;
  const grandTotal = materialsSubtotal + laborSubtotal;

  return (
    <div className="flex h-screen bg-slate-950 font-sans text-slate-100 overflow-hidden">
      {/* LEFT PANEL: File Ingestion & AI Agent Workspace */}
      <div className="w-[55%] flex flex-col justify-between border-r border-slate-800 bg-slate-900 p-6 overflow-y-auto">
        <div className="space-y-6">
          {/* Header & Logo */}
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <div className="flex items-center justify-center w-10 h-10 rounded-xl bg-emerald-500/10 border border-emerald-500/30 text-emerald-400 shadow-[0_0_15px_rgba(16,185,129,0.15)]">
                <svg className="w-6 h-6 animate-pulse" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364-6.364l-.707.707M6.343 17.657l-.707.707m0-12.728l.707.707m12.728 12.728l.707.707M12 8a4 4 0 100 8 4 4 0 000-8z" />
                </svg>
              </div>
              <div>
                <h1 className="text-lg font-bold tracking-tight text-white flex items-center">
                  Solar Estimate <span className="text-emerald-400 ml-1.5 font-semibold px-2 py-0.5 rounded bg-emerald-500/10 text-xs border border-emerald-500/20">Copilot</span>
                </h1>
                <p className="text-xs text-slate-400">Enterprise AI Solar Estimation Engine</p>
              </div>
            </div>

            {/* Help Icon Button */}
            <button
              id="help-modal-trigger"
              type="button"
              onClick={() => setHelpModalOpen(true)}
              title="System Requirements & Calculation Guide"
              className="group flex items-center justify-center w-8 h-8 rounded-full border border-slate-700 bg-slate-800/60 hover:border-emerald-500/50 hover:bg-emerald-500/10 text-slate-400 hover:text-emerald-400 transition-all duration-200 shadow-sm hover:shadow-emerald-500/10 flex-shrink-0"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M8.228 9c.549-1.165 2.03-2 3.772-2 2.21 0 4 1.343 4 3 0 1.4-1.278 2.575-3.006 2.907-.542.104-.994.54-.994 1.093m0 3h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
            </button>
          </div>

          {/* 📂 DOCUMENT INGESTION DECK */}
          <div className="relative overflow-hidden rounded-2xl border border-slate-800 bg-slate-950/70 p-5 shadow-inner backdrop-blur-md">
            <div className="absolute top-0 right-0 w-24 h-24 bg-emerald-500/5 blur-2xl rounded-full" />
            <h2 className="text-xs font-semibold text-slate-300 tracking-wider uppercase mb-4 flex items-center">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 mr-2 animate-ping" />
              1. Project Document Ingestion
            </h2>
            
            <div className="grid grid-cols-2 gap-4">
              {/* Field Study Upload */}
              <div className="group relative flex flex-col justify-center rounded-xl border border-dashed border-slate-800 hover:border-emerald-500/40 bg-slate-900/50 p-4 transition-all duration-300">
                <label className="text-xs font-medium text-slate-400 mb-2">Field Study Excel</label>
                <div className="relative">
                  <input 
                    type="file" 
                    accept=".xlsx, .xls"
                    onChange={(e) => handleFileChange(e, 'field')}
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                  />
                  <div className="flex flex-col items-center justify-center py-2 text-center">
                    <svg className="w-6 h-6 text-slate-500 group-hover:text-emerald-400 transition" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 13h6m-3-3v6m-9 1V4a2 2 0 012-2h6l2 2h6a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2z" />
                    </svg>
                    <span className="text-[11px] text-slate-400 mt-2 font-medium">Click or Drag Sheet</span>
                  </div>
                </div>

                {/* Google Drive Trigger for Field Study */}
                <button
                  type="button"
                  onClick={() => handleGoogleDrivePicker('field')}
                  className="mt-3 flex items-center justify-center space-x-2 w-full rounded-lg border border-slate-800 bg-slate-900/80 hover:bg-slate-850 hover:border-blue-500/30 px-3 py-2 text-[11px] font-medium text-slate-300 transition-all"
                >
                  <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zM19 18H6c-2.21 0-4-1.79-4-4 0-2.05 1.53-3.76 3.56-3.97l1.07-.11.5-.95C8.08 7.14 9.94 6 12 6c2.62 0 4.88 1.86 5.39 4.43l.3 1.5 1.53.11c1.56.1 2.78 1.41 2.78 2.96 0 1.65-1.35 3-3 3z" />
                  </svg>
                  <span>Import from Drive</span>
                </button>

                {fieldStudyFile && (
                  <div className="mt-3 flex items-center space-x-1.5 justify-center bg-emerald-950/30 border border-emerald-900/40 px-2 py-1 rounded text-[10px] text-emerald-400 font-mono truncate">
                    <span>📄</span>
                    <span className="truncate">{fieldStudyFile.name}</span>
                  </div>
                )}
              </div>

              {/* Price List Upload */}
              <div className="group relative flex flex-col justify-center rounded-xl border border-dashed border-slate-800 hover:border-emerald-500/40 bg-slate-900/50 p-4 transition-all duration-300">
                <label className="text-xs font-medium text-slate-400 mb-2">Material Price List</label>
                <div className="relative">
                  <input 
                    type="file" 
                    accept=".xlsx, .xls"
                    onChange={(e) => handleFileChange(e, 'price')}
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                  />
                  <div className="flex flex-col items-center justify-center py-2 text-center">
                    <svg className="w-6 h-6 text-slate-500 group-hover:text-emerald-400 transition" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                    <span className="text-[11px] text-slate-400 mt-2 font-medium">Click or Drag Index</span>
                  </div>
                </div>

                {/* Google Drive Trigger for Price List */}
                <button
                  type="button"
                  onClick={() => handleGoogleDrivePicker('price')}
                  className="mt-3 flex items-center justify-center space-x-2 w-full rounded-lg border border-slate-800 bg-slate-900/80 hover:bg-slate-850 hover:border-blue-500/30 px-3 py-2 text-[11px] font-medium text-slate-300 transition-all"
                >
                  <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96z" />
                  </svg>
                  <span>Import from Drive</span>
                </button>
                
                {priceListFile && (
                  <div className="mt-3 flex items-center space-x-1.5 justify-center bg-emerald-950/30 border border-emerald-900/40 px-2 py-1 rounded text-[10px] text-emerald-400 font-mono truncate">
                    <span>📄</span>
                    <span className="truncate">{priceListFile.name}</span>
                  </div>
                )}
              </div>
            </div>

            <button 
              onClick={handleAttachFilesToAI}
              className="mt-4 w-full bg-emerald-600 hover:bg-emerald-500 active:scale-[0.99] text-white text-xs font-semibold py-2.5 rounded-lg shadow-lg hover:shadow-emerald-500/20 transition-all duration-300"
            >
              Analyze & Sync Files to AI Agent
            </button>
            
            {uploadStatus && (
              <p className={`text-xs font-medium text-center mt-3 p-2 rounded ${
                uploadStatus.includes('❌') ? 'bg-rose-950/40 text-rose-300 border border-rose-900/40' : 'bg-slate-900 text-slate-300 border border-slate-800'
              }`}>
                {uploadStatus}
              </p>
            )}
          </div>

          <div className="border-b border-slate-850" />
          
        </div>
      </div>

      {/* RIGHT PANEL: Dynamic Ingestion Review & Excel Compilation Engine */}
      <div className="w-[45%] flex flex-col bg-slate-950 p-6 overflow-y-auto">
        {calculationResult ? (
          <div className="flex flex-col h-full justify-between space-y-6">
            
            {/* COMPILATION METRICS SECTION */}
            <div className="space-y-6">
              <div className="flex justify-between items-center">
                <h3 className="text-sm font-bold tracking-wider text-slate-200 uppercase flex items-center">
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-400 mr-2 shadow-[0_0_10px_rgba(52,211,153,0.5)]" />
                  {calculationResult.customerName ? `${calculationResult.customerName} — BOQ` : 'Calculation Approvals (Preview)'}
                </h3>
                <span className="text-[10px] font-mono text-emerald-400 border border-emerald-500/20 px-2 py-0.5 rounded bg-emerald-500/5">
                  {calculationResult.metrics ? `${calculationResult.metrics.apartmentCount} Apts` : '100% Verified'}
                </span>
              </div>

              {/* Statistical Ingest KPI Grid */}
              <div className="grid grid-cols-3 gap-3">
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 space-y-1">
                  <span className="text-[10px] font-medium text-slate-500 uppercase tracking-wider block">Materials</span>
                  <div className="flex items-baseline space-x-1">
                    <span className="text-lg font-bold text-white">{calculationResult.materials?.length || 0}</span>
                    <span className="text-[10px] text-slate-400 font-medium">SKUs</span>
                  </div>
                </div>
                
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 space-y-1">
                  <span className="text-[10px] font-medium text-slate-500 uppercase tracking-wider block">Labor Hours</span>
                  <div className="flex items-baseline space-x-1">
                    <span className="text-lg font-bold text-white">
                      {calculationResult.labor?.reduce((acc: number, l: any) => acc + l.hours, 0) || 0}
                    </span>
                    <span className="text-[10px] text-slate-400 font-medium">hrs</span>
                  </div>
                </div>

                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 space-y-1">
                  <span className="text-[10px] font-medium text-emerald-400 uppercase tracking-wider block">Peak System Load</span>
                  <div className="flex items-baseline space-x-1">
                    <span className="text-lg font-bold text-emerald-400">
                      {calculationResult.metrics?.peakKW?.toFixed(1) || '0.0'}
                    </span>
                    <span className="text-[10px] text-emerald-400 font-medium">kW</span>
                  </div>
                </div>
              </div>

              {/* Detailed Itemized Materials Breakdown */}
              <div className="space-y-3">
                <h4 className="text-xs font-semibold text-slate-400 tracking-wide uppercase">Itemized Pricing Breakdown</h4>
                
                <div className="rounded-xl border border-slate-850 bg-slate-900/30 overflow-hidden">
                  <div className="max-h-[30vh] overflow-y-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-900/90 text-[10px] uppercase text-slate-500 border-b border-slate-850 sticky top-0 z-10">
                          <th className="py-2.5 px-3 font-semibold">Component Description</th>
                          <th className="py-2.5 px-3 font-semibold text-right">Qty</th>
                          <th className="py-2.5 px-3 font-semibold text-right">Unit Rate</th>
                          <th className="py-2.5 px-3 font-semibold text-right">Total</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-850 font-mono text-[11px]">
                        {/* Materials Section */}
                        <tr className="bg-slate-950/40">
                          <td colSpan={4} className="py-1.5 px-3 text-[10px] font-bold text-emerald-500/80 tracking-wide uppercase">
                            📦 Solar Materials
                          </td>
                        </tr>
                        {calculationResult.materials?.map((m: any, idx: number) => (
                          <tr key={idx} className="hover:bg-slate-800/30 text-slate-300">
                            <td className="py-2 px-3 truncate max-w-[160px] text-left">{m.name}</td>
                            <td className="py-2 px-3 text-right">{m.quantity}</td>
                            <td className="py-2 px-3 text-right">{m.unitPrice.toLocaleString()} XAF</td>
                            <td className="py-2 px-3 text-right text-emerald-400/90">{(m.quantity * m.unitPrice).toLocaleString()} XAF</td>
                          </tr>
                        ))}

                        {/* Labor Section */}
                        <tr className="bg-slate-950/40">
                          <td colSpan={4} className="py-1.5 px-3 text-[10px] font-bold text-cyan-500/80 tracking-wide uppercase border-t border-slate-850">
                            🛠️ Technical Labor
                          </td>
                        </tr>
                        {calculationResult.labor?.map((l: any, idx: number) => (
                          <tr key={idx} className="hover:bg-slate-800/30 text-slate-300">
                            <td className="py-2 px-3 truncate max-w-[160px] text-left">{l.description}</td>
                            <td className="py-2 px-3 text-right text-slate-500 italic text-[10px]">30% of Equipment</td>
                            <td className="py-2 px-3 text-right text-slate-500 italic text-[10px]">Flat Rate</td>
                            <td className="py-2 px-3 text-right text-emerald-400/90">{l.hourlyRate.toLocaleString()} XAF</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>

              {/* Solar Metrics Banner - only shown when parsed from ingestion engine */}
              {calculationResult.metrics && (
                <div className="grid grid-cols-3 gap-2 mb-1">
                  <div className="bg-slate-900/80 border border-slate-800 rounded-lg p-2.5 text-center">
                    <span className="text-[9px] text-slate-500 uppercase tracking-wider block">Peak Load</span>
                    <span className="text-sm font-bold text-amber-400">{calculationResult.metrics.peakKW?.toFixed(1)} kW</span>
                  </div>
                  <div className="bg-slate-900/80 border border-slate-800 rounded-lg p-2.5 text-center">
                    <span className="text-[9px] text-slate-500 uppercase tracking-wider block">Day kWh</span>
                    <span className="text-sm font-bold text-sky-400">{calculationResult.metrics.dayConsumptionKWh?.toFixed(1)}</span>
                  </div>
                  <div className="bg-slate-900/80 border border-slate-800 rounded-lg p-2.5 text-center">
                    <span className="text-[9px] text-slate-500 uppercase tracking-wider block">Night kWh</span>
                    <span className="text-sm font-bold text-violet-400">{calculationResult.metrics.nightConsumptionKWh?.toFixed(1)}</span>
                  </div>
                </div>
              )}

              {/* Glowing Grand Totals Board */}
              <div className="rounded-xl border border-emerald-500/20 bg-gradient-to-br from-slate-900 to-emerald-950/20 p-4 space-y-2.5 shadow-md shadow-emerald-950/10">
                <div className="flex justify-between items-center text-xs">
                  <span className="text-slate-400">Materials Subtotal</span>
                  <span className="font-mono text-slate-300">{materialsSubtotal.toLocaleString()} XAF</span>
                </div>
                <div className="flex justify-between items-center text-xs">
                  <span className="text-slate-400">Labor Subtotal</span>
                  <span className="font-mono text-slate-300">{laborSubtotal.toLocaleString()} XAF</span>
                </div>
                <div className="border-t border-slate-800 my-2" />
                <div className="flex justify-between items-center">
                  <div>
                    <span className="text-xs font-semibold text-slate-300 uppercase block">Grand Total</span>
                    <span className="text-[10px] text-slate-500 font-medium">Direct Materials & Labor</span>
                  </div>
                  <span className="text-xl font-bold font-mono text-emerald-400 shadow-emerald-400/10">
                    {grandTotal.toLocaleString(undefined, { maximumFractionDigits: 0 })} XAF
                  </span>
                </div>
              </div>
            </div>

            {/* DOWNLOAD EXCEL CALL TO ACTION */}
            <div className="pt-4 border-t border-slate-850 space-y-3">
              <div className="flex items-center space-x-2 text-[10px] text-slate-400 bg-slate-900/60 p-2.5 rounded-lg border border-slate-850">
                <svg className="w-4 h-4 text-emerald-400 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
                <span>Estimated values compiled cleanly. Click below to download the stylized XLSX deliverable complete with auto-generated formulas.</span>
              </div>
              <button 
                onClick={downloadExcel} 
                className="w-full bg-gradient-to-r from-emerald-600 to-emerald-500 hover:from-emerald-500 hover:to-emerald-400 active:scale-[0.99] text-white py-3 rounded-xl font-bold text-xs tracking-wider uppercase transition-all duration-300 shadow-[0_4px_20px_rgba(16,185,129,0.15)] flex items-center justify-center space-x-2"
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                <span>Compile & Download XLSX Report</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="flex-1 flex flex-col justify-center items-center text-center p-8">
            {/* Interactive Dashboard Waiting Visual */}
            <div className="relative mb-6">
              <div className="absolute inset-0 bg-emerald-500/10 blur-3xl rounded-full animate-pulse" />
              <div className="relative flex items-center justify-center w-24 h-24 rounded-3xl bg-slate-900 border border-slate-800 shadow-2xl">
                <svg className="w-10 h-10 text-slate-700 animate-spin" style={{ animationDuration: '6s' }} fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                </svg>
              </div>
            </div>
            
            <h3 className="text-base font-bold text-slate-350 tracking-wide uppercase">Math Compiler Dormant</h3>
            <p className="text-xs text-slate-500 mt-2 max-w-sm leading-relaxed">
              Upload your site Field Study and warehouse Price List spreadsheet arrays. Once synced, our AI agent will map parameters to execute precision pricing estimates instantly.
            </p>
            
            {/* Visual workflow step helper */}
            <div className="mt-8 flex items-center space-x-2 text-[10px] text-slate-650 bg-slate-900/30 px-4 py-2.5 rounded-xl border border-slate-900">
              <span>Ingest Sheets</span>
              <span>➡️</span>
              <span>Synchronize Context</span>
              <span>➡️</span>
              <span className="text-emerald-500/70 font-semibold">Compile Estimates</span>
            </div>
          </div>
        )}
      </div>

      {/* 🔮 GOOGLE DRIVE IMPORT GLASSMORPHIC MODAL */}
      {driveModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/80 backdrop-blur-sm p-4">
          <div className="relative w-full max-w-md overflow-hidden rounded-2xl border border-slate-800 bg-slate-900/95 p-6 shadow-2xl transition-all duration-300">
            {/* Background ambient glow */}
            <div className="absolute -top-12 -right-12 w-28 h-28 bg-emerald-500/10 blur-3xl rounded-full" />
            <div className="absolute -bottom-12 -left-12 w-28 h-28 bg-blue-500/10 blur-3xl rounded-full" />

            <div className="flex items-center space-x-3 mb-4">
              <div className="flex items-center justify-center w-8 h-8 rounded-lg bg-emerald-500/10 border border-emerald-500/30 text-emerald-400">
                <svg className="w-5 h-5 animate-pulse" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96z" />
                </svg>
              </div>
              <h3 className="text-sm font-bold text-white uppercase tracking-wider">
                Import {driveImportType === 'field' ? 'Field Study' : 'Price List'} from Drive
              </h3>
            </div>

            <p className="text-xs text-slate-400 leading-relaxed mb-4">
              Paste the Google Drive or Google Sheets sharing link below. Ensure the file has sharing settings configured as <span className="text-emerald-400 font-semibold">"Anyone with the link can view"</span> so the parsing engine can fetch it.
            </p>

            <div className="space-y-4">
              <div>
                <label className="text-[10px] font-bold uppercase tracking-wider text-slate-500 block mb-1.5">
                  Google Drive Shareable Link
                </label>
                <input
                  type="text"
                  placeholder="https://drive.google.com/file/d/.../view"
                  value={driveUrlInput}
                  onChange={(e) => setDriveUrlInput(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 focus:border-emerald-500/50 rounded-xl px-3 py-2 text-xs text-slate-100 placeholder-slate-600 outline-none transition-all"
                  disabled={driveLoading}
                />
              </div>

              {driveError && (
                <div className="bg-rose-950/40 border border-rose-900/40 text-rose-300 text-xs p-2.5 rounded-lg font-medium leading-relaxed">
                  {driveError}
                </div>
              )}

              <div className="flex items-center space-x-3 pt-2">
                <button
                  type="button"
                  onClick={() => setDriveModalOpen(false)}
                  className="flex-1 bg-slate-800 hover:bg-slate-750 active:scale-[0.99] text-slate-350 text-xs font-semibold py-2.5 rounded-lg border border-slate-750 transition-all duration-200"
                  disabled={driveLoading}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleImportFromDrive}
                  className="flex-1 bg-emerald-600 hover:bg-emerald-500 active:scale-[0.99] text-white text-xs font-bold py-2.5 rounded-lg shadow-lg shadow-emerald-950/20 transition-all duration-200 flex items-center justify-center space-x-1.5"
                  disabled={driveLoading}
                >
                  {driveLoading ? (
                    <>
                      <svg className="animate-spin h-3.5 w-3.5 text-white" fill="none" viewBox="0 0 24 24">
                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                      </svg>
                      <span>Fetching...</span>
                    </>
                  ) : (
                    <span>Import File</span>
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 📖 HELP MODAL — System Requirements & Calculation Guide */}
      {helpModalOpen && (
        <div className="fixed inset-0 z-50 flex items-start justify-end bg-slate-950/70 backdrop-blur-sm p-4" onClick={() => setHelpModalOpen(false)}>
          <div
            className="relative w-full max-w-lg max-h-[92vh] overflow-y-auto rounded-2xl border border-slate-700/70 bg-slate-900/98 shadow-2xl shadow-slate-950/60 transition-all duration-300"
            onClick={(e) => e.stopPropagation()}
          >
            {/* Ambient glows */}
            <div className="absolute -top-10 -right-10 w-36 h-36 bg-emerald-500/8 blur-3xl rounded-full pointer-events-none" />
            <div className="absolute -bottom-10 -left-10 w-36 h-36 bg-blue-500/8 blur-3xl rounded-full pointer-events-none" />

            {/* Header */}
            <div className="sticky top-0 z-10 flex items-center justify-between px-5 py-4 border-b border-slate-800 bg-slate-900/95 backdrop-blur-md rounded-t-2xl">
              <div className="flex items-center space-x-2.5">
                <div className="flex items-center justify-center w-7 h-7 rounded-lg bg-emerald-500/10 border border-emerald-500/30 text-emerald-400">
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M8.228 9c.549-1.165 2.03-2 3.772-2 2.21 0 4 1.343 4 3 0 1.4-1.278 2.575-3.006 2.907-.542.104-.994.54-.994 1.093m0 3h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                </div>
                <div>
                  <h2 className="text-sm font-bold text-white tracking-wide">System Guide</h2>
                  <p className="text-[10px] text-slate-500">Requirements & Calculation Engine Reference</p>
                </div>
              </div>
              <button onClick={() => setHelpModalOpen(false)} className="text-slate-500 hover:text-white transition-colors p-1 rounded-lg hover:bg-slate-800">
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
              </button>
            </div>

            <div className="px-5 py-4 space-y-5 text-xs">

              {/* ── FILE 1: Field Study ── */}
              <section>
                <div className="flex items-center space-x-2 mb-2.5">
                  <span className="text-base">📋</span>
                  <h3 className="text-xs font-bold text-slate-200 uppercase tracking-widest">File 1 — Field Study Workbook</h3>
                </div>
                <div className="rounded-xl border border-slate-800 bg-slate-950/50 divide-y divide-slate-800 overflow-hidden">
                  <div className="flex items-start px-3 py-2.5 space-x-3">
                    <span className="text-emerald-400 font-bold mt-0.5 flex-shrink-0">TYPE</span>
                    <span className="text-slate-300">.xlsx or .xls (Microsoft Excel Workbook)</span>
                  </div>
                  <div className="flex items-start px-3 py-2.5 space-x-3">
                    <span className="text-emerald-400 font-bold mt-0.5 flex-shrink-0">SHEETS</span>
                    <span className="text-slate-300">One or more sheets named exactly as <code className="bg-slate-800 px-1 rounded text-emerald-300">01_Comsumption_profile</code>, <code className="bg-slate-800 px-1 rounded text-emerald-300">02_Comsumption_profile</code>, etc. (sequential two-digit prefix)</span>
                  </div>
                  <div className="flex items-start px-3 py-2.5 space-x-3">
                    <span className="text-emerald-400 font-bold mt-0.5 flex-shrink-0">A1</span>
                    <span className="text-slate-300">Must contain the customer name in the format: <code className="bg-slate-800 px-1 rounded text-emerald-300">CUSTOMER'S NAME: [Name]</code></span>
                  </div>
                  <div className="flex items-start px-3 py-2.5 space-x-3">
                    <span className="text-emerald-400 font-bold mt-0.5 flex-shrink-0">ROW 6</span>
                    <span className="text-slate-300">Column headers row (ignored by parser)</span>
                  </div>
                  <div className="px-3 py-2.5 space-y-1.5">
                    <span className="text-emerald-400 font-bold block">ROW 7+ — Device Data Columns:</span>
                    <div className="grid grid-cols-2 gap-1 mt-1">
                      {[
                        ['A', 'Device Name'],
                        ['B', 'Power (W)'],
                        ['C', 'Voltage (V)'],
                        ['D', 'Quantity'],
                        ['E', 'Day Runtime (hrs)'],
                        ['F', 'Night Runtime (hrs)'],
                        ['G', 'Day Consumption (Wh)'],
                        ['H', 'Night Consumption (Wh)'],
                      ].map(([col, label]) => (
                        <div key={col} className="flex items-center space-x-1.5 bg-slate-900 rounded px-2 py-1">
                          <span className="text-[10px] font-bold text-emerald-400 w-4 flex-shrink-0">{col}</span>
                          <span className="text-slate-400">{label}</span>
                        </div>
                      ))}
                    </div>
                    <p className="text-slate-500 mt-1.5 italic">Tip: If G or H columns are blank, they are auto-calculated as B × D × E (or F).</p>
                  </div>
                </div>
              </section>

              {/* ── FILE 2: Price List ── */}
              <section>
                <div className="flex items-center space-x-2 mb-2.5">
                  <span className="text-base">💰</span>
                  <h3 className="text-xs font-bold text-slate-200 uppercase tracking-widest">File 2 — Material Price List Workbook</h3>
                </div>
                <div className="rounded-xl border border-slate-800 bg-slate-950/50 divide-y divide-slate-800 overflow-hidden">
                  <div className="flex items-start px-3 py-2.5 space-x-3">
                    <span className="text-emerald-400 font-bold mt-0.5 flex-shrink-0">TYPE</span>
                    <span className="text-slate-300">.xlsx or .xls (Microsoft Excel Workbook)</span>
                  </div>
                  <div className="px-3 py-2.5 space-y-2">
                    <span className="text-emerald-400 font-bold block">REQUIRED SHEETS (4 total):</span>
                    {[
                      { name: 'Solar inverters', color: 'bg-amber-500/10 border-amber-500/30 text-amber-400' },
                      { name: 'Solar batteries', color: 'bg-blue-500/10 border-blue-500/30 text-blue-400' },
                      { name: 'Solar Panels',    color: 'bg-emerald-500/10 border-emerald-500/30 text-emerald-400' },
                      { name: 'Cables',          color: 'bg-violet-500/10 border-violet-500/30 text-violet-400' },
                    ].map(s => (
                      <div key={s.name} className={`inline-flex items-center px-2 py-0.5 rounded border text-[10px] font-mono mr-1.5 ${s.color}`}>{s.name}</div>
                    ))}
                  </div>
                  <div className="px-3 py-2.5 space-y-1">
                    <span className="text-emerald-400 font-bold block mb-1.5">COLUMN HEADERS (per sheet):</span>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-[10px]">
                        <thead>
                          <tr className="text-slate-500 border-b border-slate-800">
                            <th className="pb-1.5 pr-3 font-semibold">Header</th>
                            <th className="pb-1.5 pr-3 font-semibold">Accepts / Aliases</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800/60">
                          {[
                            ['S/N', 'Row index — used as end-of-table sentinel'],
                            ['Type', 'Model name / Product description / SKU'],
                            ['Power Rating', 'kW for inverters · kWh for batteries · W for panels'],
                            ['Brand', 'Manufacturer name'],
                            ['Price', 'Unit cost in XAF (also accepts: cost, rate, xaf)'],
                          ].map(([h, a]) => (
                            <tr key={h}>
                              <td className="py-1.5 pr-3 font-mono text-emerald-300 font-semibold">{h}</td>
                              <td className="py-1.5 text-slate-400">{a}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <p className="text-slate-500 mt-2 italic">Headers can appear in any row within the first 15 rows. If a sheet has no parseable data, the engine uses a &ldquo;—&rdquo; placeholder.</p>
                  </div>
                </div>
              </section>

              {/* ── Calculation Engine ── */}
              <section>
                <div className="flex items-center space-x-2 mb-2.5">
                  <span className="text-base">⚙️</span>
                  <h3 className="text-xs font-bold text-slate-200 uppercase tracking-widest">Calculation Engine Reference</h3>
                </div>
                <div className="space-y-2">

                  {/* Step 1 */}
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 px-3 py-2.5">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">① Load Aggregation (from Field Study)</p>
                    <div className="space-y-1 font-mono text-[10px]">
                      <div className="flex justify-between"><span className="text-slate-400">Peak Load (kW)</span><span className="text-emerald-300">= Σ (Power × Qty) ÷ 1000</span></div>
                      <div className="flex justify-between"><span className="text-slate-400">Day Consumption (kWh)</span><span className="text-emerald-300">= Σ (Power × Qty × Day hrs) ÷ 1000</span></div>
                      <div className="flex justify-between"><span className="text-slate-400">Night Consumption (kWh)</span><span className="text-emerald-300">= Σ (Power × Qty × Night hrs) ÷ 1000</span></div>
                    </div>
                  </div>

                  {/* Step 2 */}
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 px-3 py-2.5">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">② Inverter Selection</p>
                    <p className="text-slate-400 leading-relaxed">Finds the <span className="text-white font-semibold">smallest inverter</span> in the price list whose <span className="text-emerald-400">powerKW ≥ Peak Load</span>. If no single unit is large enough, the engine stacks the largest available inverter in multiples and raises a <span className="text-amber-400">CAPACITY_OVERFLOW_WARNING</span>.</p>
                  </div>

                  {/* Step 3 */}
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 px-3 py-2.5">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">③ Battery Sizing</p>
                    <div className="space-y-1 font-mono text-[10px] mb-1.5">
                      <div className="flex justify-between"><span className="text-slate-400">Target Capacity</span><span className="text-emerald-300">= Night kWh × 1.20 (20% safety buffer)</span></div>
                    </div>
                    <p className="text-slate-400 leading-relaxed">Selects the <span className="text-white font-semibold">smallest battery</span> covering the target, prioritising <span className="text-blue-400">Lithium</span> over Gel. If no single unit covers it, stacks multiples in parallel.</p>
                  </div>

                  {/* Step 4 */}
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 px-3 py-2.5">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">④ Solar Panel Array</p>
                    <div className="space-y-1 font-mono text-[10px] mb-1.5">
                      <div className="flex justify-between"><span className="text-slate-400">Daily Output / panel</span><span className="text-emerald-300">= Power(W) × 5 peak sun hrs ÷ 1000</span></div>
                      <div className="flex justify-between"><span className="text-slate-400">Panel Count</span><span className="text-emerald-300">= ⌈ Day kWh ÷ Daily Output ⌉</span></div>
                    </div>
                    <p className="text-slate-400 leading-relaxed">Always picks the <span className="text-white font-semibold">highest-efficiency panel</span> in the price list to minimise roof footprint. A warning fires if the array count exceeds 80 panels.</p>
                  </div>

                  {/* Step 5 */}
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 px-3 py-2.5">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">⑤ Cable Sizing</p>
                    <div className="space-y-1 font-mono text-[10px] mb-1.5">
                      <div className="flex justify-between"><span className="text-slate-400">Total Amperage</span><span className="text-emerald-300">= Inverter A × Inverter Qty</span></div>
                      <div className="flex justify-between"><span className="text-slate-400">Cable Metres</span><span className="text-emerald-300">= 50 + (Panel Count × 2)</span></div>
                    </div>
                    <p className="text-slate-400 leading-relaxed">Selects the <span className="text-white font-semibold">cheapest cable</span> whose maxAmperage covers total amperage. If no amperage column exists, rating is inferred from the cable cross-section (mm²).</p>
                  </div>

                  {/* Step 6 */}
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 px-3 py-2.5">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">⑥ Labor & Grand Total</p>
                    <div className="space-y-1 font-mono text-[10px]">
                      <div className="flex justify-between"><span className="text-slate-400">Labor Cost</span><span className="text-emerald-300">= Materials Subtotal × 30%</span></div>
                      <div className="flex justify-between font-bold"><span className="text-white">Grand Total</span><span className="text-emerald-400">= Materials + Labor</span></div>
                    </div>
                  </div>

                </div>
              </section>

              {/* Footer */}
              <div className="flex items-center space-x-2 text-[10px] text-slate-600 border-t border-slate-800 pt-3">
                <svg className="w-3.5 h-3.5 text-emerald-600 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                <span>All monetary values are in <strong className="text-slate-400">XAF (Central African Franc)</strong>. Upload new price list to update rates at any time without code changes.</span>
              </div>

            </div>
          </div>
        </div>
      )}
    </div>
  );
}