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

  // Initializing useChat with DefaultChatTransport for custom endpoint and onToolCall handler
  const { messages, sendMessage, setMessages, status } = useChat({
    transport: new DefaultChatTransport({ api: '/api/chat' }),
    onToolCall({ toolCall }) {
      if (toolCall.toolName === 'calculateEstimate') {
        // In AI SDK v6, the arguments are passed as `input` instead of `args`
        setCalculationResult(toolCall.input as any);
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
        metrics: parsed.metrics
      });

      // Format a detailed system notice so the AI agent is aware of all computed values
      const m = parsed.metrics;
      const warningsBlock = parsed.warnings?.length
        ? `\n⚠️ SYSTEM WARNINGS:\n${parsed.warnings.map((w: string) => `  • ${w}`).join('\n')}`
        : '';

      const totalMat = parsed.materials?.reduce((acc: number, item: any) => acc + item.quantity * item.unitPrice, 0) || 0;
      const totalLab = parsed.labor?.reduce((acc: number, item: any) => acc + item.hours * item.hourlyRate, 0) || 0;
      const grandTotalXAF = (totalMat + totalLab) * (1 + (parsed.marginPercentage || 15) / 100);

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
  • Margin: ${parsed.marginPercentage}%
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
  const grandTotal = (materialsSubtotal + laborSubtotal) * (1 + (calculationResult?.marginPercentage || 15) / 100);

  return (
    <div className="flex h-screen bg-slate-950 font-sans text-slate-100 overflow-hidden">
      {/* LEFT PANEL: File Ingestion & AI Agent Workspace */}
      <div className="w-[55%] flex flex-col justify-between border-r border-slate-800 bg-slate-900 p-6 overflow-y-auto">
        <div className="space-y-6">
          {/* Header & Logo */}
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

                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 space-y-1 animate-pulse">
                  <span className="text-[10px] font-medium text-emerald-500 uppercase tracking-wider block">Margin Rate</span>
                  <div className="flex items-baseline space-x-1">
                    <span className="text-lg font-bold text-emerald-400">{calculationResult.marginPercentage || 15}</span>
                    <span className="text-[10px] text-emerald-500 font-medium">%</span>
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
                          <th className="py-2.5 px-3 font-semibold text-right">Qty/Hrs</th>
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
                    <span className="text-[10px] text-emerald-400 font-medium">Including {calculationResult.marginPercentage || 15}% Markup</span>
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
    </div>
  );
}