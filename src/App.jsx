import React, { useState, useEffect, useRef } from 'react';
import { 
  FileText, 
  Settings, 
  Send, 
  Download, 
  FileCheck, 
  BrainCircuit, 
  Loader2, 
  AlertCircle, 
  Plus, 
  Trash2, 
  ChevronRight, 
  HelpCircle,
  FolderOpen
} from 'lucide-react';
import { renderAsync } from 'docx-preview';

export default function App() {
  // App States
  const [loading, setLoading] = useState(false);
  const [statusMsg, setStatusMsg] = useState('');
  
  // Files States
  const [coursePlanFile, setCoursePlanFile] = useState(null);
  const [coursePlanText, setCoursePlanText] = useState('');
  const [coursePlanStatus, setCoursePlanStatus] = useState('pending'); // pending, reading, ready
  
  const [mesepBookFile, setMesepBookFile] = useState(null);
  const [mesepBookText, setMesepBookText] = useState('');
  const [mesepBookStatus, setMesepBookStatus] = useState('pending'); // pending, reading, ready
  
  const [templateFile, setTemplateFile] = useState(null);
  const [templateStatus, setTemplateStatus] = useState('pending'); // pending, ready

  // Chat States
  const [chatMessage, setChatMessage] = useState('');
  const [chatHistory, setChatHistory] = useState([]);
  const [chatLoading, setChatLoading] = useState(false);
  const chatBottomRef = useRef(null);

  // Teaching Plan State (JSON from Gemini)
  const [planData, setPlanData] = useState(null);
  const [activeSaTab, setActiveSaTab] = useState(0);
  const [generatedDocxBlob, setGeneratedDocxBlob] = useState(null);

  // Scroll chat to bottom
  useEffect(() => {
    chatBottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [chatHistory]);

  // Extract Text from PDF/DOCX
  const handleFileUpload = async (e, type) => {
    const file = e.target.files[0];
    if (!file) return;

    if (type === 'course_plan') {
      setCoursePlanFile(file);
      setCoursePlanStatus('reading');
      const text = await extractTextFromFile(file);
      if (text) {
        setCoursePlanText(text);
        setCoursePlanStatus('ready');
      } else {
        setCoursePlanStatus('pending');
      }
    } else if (type === 'mesep_book') {
      setMesepBookFile(file);
      setMesepBookStatus('reading');
      const text = await extractTextFromFile(file);
      if (text) {
        setMesepBookText(text);
        setMesepBookStatus('ready');
      } else {
        setMesepBookStatus('pending');
      }
    } else if (type === 'template') {
      setTemplateFile(file);
      setTemplateStatus('ready');
    }
  };

  const extractTextFromFile = async (file) => {
    const formData = new FormData();
    formData.append('file', file);
    try {
      const response = await fetch('/api/extract', {
        method: 'POST',
        body: formData,
      });
      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.detail || 'Erro ao extrair arquivo');
      }
      const data = await response.json();
      return data.text;
    } catch (e) {
      alert(`Falha ao ler o arquivo: ${e.message}`);
      return '';
    }
  };

  // NotebookLM Chat
  const sendChatMessage = async (e) => {
    e?.preventDefault();
    if (!chatMessage.trim() || chatLoading) return;

    const userMsg = chatMessage.trim();
    setChatMessage('');
    setChatHistory(prev => [...prev, { role: 'user', text: userMsg }]);
    setChatLoading(true);

    try {
      const formData = new FormData();
      formData.append('message', userMsg);
      formData.append('history', JSON.stringify(chatHistory));
      formData.append('course_plan', coursePlanText);
      formData.append('mesep_book', mesepBookText);

      const response = await fetch('/api/chat', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.detail || 'Erro ao conversar');
      }

      const data = await response.json();
      setChatHistory(prev => [...prev, { role: 'model', text: data.response }]);
    } catch (e) {
      setChatHistory(prev => [...prev, { role: 'model', text: `Erro: ${e.message}` }]);
    } finally {
      setChatLoading(false);
    }
  };

  // Suggestion chips
  const handleSuggestionClick = (msg) => {
    setChatMessage(msg);
  };

  // Analyze course plan and generate SAs
  const analyzeCoursePlan = async () => {
    if (!coursePlanText) {
      alert('Por favor, faça o upload do Plano de Curso.');
      return;
    }
    if (!templateFile) {
      alert('Por favor, faça o upload do Modelo DOCX.');
      return;
    }

    setLoading(true);
    setStatusMsg('Analisando o Plano de Curso com IA...');
    try {
      const formData = new FormData();
      formData.append('course_plan', coursePlanText);
      formData.append('mesep_book', mesepBookText);

      const response = await fetch('/api/analyze', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.detail || 'Erro na análise do plano');
      }

      const data = await response.json();
      setPlanData(data);
      setActiveSaTab(0);
      setGeneratedDocxBlob(null);
    } catch (e) {
      alert(`Falha na análise: ${e.message}`);
    } finally {
      setLoading(false);
      setStatusMsg('');
    }
  };

  // SA modification helpers
  const updatePlanField = (field, value) => {
    setPlanData(prev => ({
      ...prev,
      [field]: value
    }));
    setGeneratedDocxBlob(null);
  };

  const updateSaField = (saIdx, field, value) => {
    setPlanData(prev => {
      const newSas = [...prev.situacoes_aprendizagem];
      newSas[saIdx] = {
        ...newSas[saIdx],
        [field]: value
      };
      return {
        ...prev,
        situacoes_aprendizagem: newSas
      };
    });
    setGeneratedDocxBlob(null);
  };

  const addSa = () => {
    if (planData.situacoes_aprendizagem.length >= 4) {
      alert('O template oficial suporta um máximo de 4 Situações de Aprendizagem.');
      return;
    }
    const newSaNum = planData.situacoes_aprendizagem.length + 1;
    const newSa = {
      titulo: `SITUAÇÃO DE APRENDIZAGEM 0${newSaNum}`,
      carga_horaria: '20h',
      estrategia_tipo: 'Situação-problema',
      capacidades_tecnicas: ['Nova capacidade técnica'],
      capacidades_socioemocionais: ['Nova capacidade socioemocional'],
      conhecimentos: ['Novo conhecimento'],
      contextualizacao: 'Cenário profissional...',
      observacoes_docente: 'Instruções para o professor...',
      desafio: 'Desafio para o aluno...',
      resultados_esperados: ['Produto final'],
      anexos: 'Figuras...',
      referencias: 'Livros...',
      instrumento_registro: [
        { capacidade: 'Nova capacidade', criterio: '[CRÍTICO] Pergunta avaliativa?' }
      ]
    };
    setPlanData(prev => ({
      ...prev,
      situacoes_aprendizagem: [...prev.situacoes_aprendizagem, newSa]
    }));
    setActiveSaTab(planData.situacoes_aprendizagem.length);
    setGeneratedDocxBlob(null);
  };

  const removeSa = (saIdx) => {
    if (planData.situacoes_aprendizagem.length <= 1) {
      alert('Você precisa ter pelo menos 1 Situação de Aprendizagem.');
      return;
    }
    if (!confirm('Deseja realmente remover esta Situação de Aprendizagem?')) return;
    
    setPlanData(prev => {
      const newSas = prev.situacoes_aprendizagem.filter((_, idx) => idx !== saIdx)
        .map((sa, idx) => ({
          ...sa,
          titulo: `SITUAÇÃO DE APRENDIZAGEM 0${idx + 1}`
        }));
      return {
        ...prev,
        situacoes_aprendizagem: newSas
      };
    });
    setActiveSaTab(0);
    setGeneratedDocxBlob(null);
  };

  // Edit list arrays inside SA
  const handleListChange = (saIdx, field, listIdx, value) => {
    setPlanData(prev => {
      const newSas = [...prev.situacoes_aprendizagem];
      const newList = [...newSas[saIdx][field]];
      newList[listIdx] = value;
      newSas[saIdx] = {
        ...newSas[saIdx],
        [field]: newList
      };
      return {
        ...prev,
        situacoes_aprendizagem: newSas
      };
    });
    setGeneratedDocxBlob(null);
  };

  const addListItem = (saIdx, field, defaultValue = '') => {
    setPlanData(prev => {
      const newSas = [...prev.situacoes_aprendizagem];
      newSas[saIdx] = {
        ...newSas[saIdx],
        [field]: [...newSas[saIdx][field], defaultValue]
      };
      return {
        ...prev,
        situacoes_aprendizagem: newSas
      };
    });
    setGeneratedDocxBlob(null);
  };

  const removeListItem = (saIdx, field, listIdx) => {
    setPlanData(prev => {
      const newSas = [...prev.situacoes_aprendizagem];
      newSas[saIdx] = {
        ...newSas[saIdx],
        [field]: newSas[saIdx][field].filter((_, idx) => idx !== listIdx)
      };
      return {
        ...prev,
        situacoes_aprendizagem: newSas
      };
    });
    setGeneratedDocxBlob(null);
  };

  // Edit Instrument of evaluation inside SA
  const handleInstrumentChange = (saIdx, instIdx, key, value) => {
    setPlanData(prev => {
      const newSas = [...prev.situacoes_aprendizagem];
      const newInst = [...newSas[saIdx].instrumento_registro];
      newInst[instIdx] = {
        ...newInst[instIdx],
        [key]: value
      };
      newSas[saIdx] = {
        ...newSas[saIdx],
        instrumento_registro: newInst
      };
      return {
        ...prev,
        situacoes_aprendizagem: newSas
      };
    });
    setGeneratedDocxBlob(null);
  };

  const addInstrumentRow = (saIdx) => {
    setPlanData(prev => {
      const newSas = [...prev.situacoes_aprendizagem];
      newSas[saIdx] = {
        ...newSas[saIdx],
        instrumento_registro: [
          ...newSas[saIdx].instrumento_registro,
          { capacidade: 'Defina a capacidade', criterio: '[CRÍTICO] Defina a pergunta do critério' }
        ]
      };
      return {
        ...prev,
        situacoes_aprendizagem: newSas
      };
    });
    setGeneratedDocxBlob(null);
  };

  const removeInstrumentRow = (saIdx, instIdx) => {
    setPlanData(prev => {
      const newSas = [...prev.situacoes_aprendizagem];
      newSas[saIdx] = {
        ...newSas[saIdx],
        instrumento_registro: newSas[saIdx].instrumento_registro.filter((_, idx) => idx !== instIdx)
      };
      return {
        ...prev,
        situacoes_aprendizagem: newSas
      };
    });
    setGeneratedDocxBlob(null);
  };

  // Generate Filled DOCX file
  const generateDocx = async () => {
    if (!templateFile) {
      alert('Por favor, envie o arquivo DOCX de modelo.');
      return;
    }
    
    setLoading(true);
    setStatusMsg('Preenchendo o arquivo DOCX...');
    
    try {
      const formData = new FormData();
      formData.append('template', templateFile);
      formData.append('data', JSON.stringify(planData));

      const response = await fetch('/api/generate', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.detail || 'Erro ao gerar arquivo');
      }

      const blob = await response.blob();
      setGeneratedDocxBlob(blob);
      return blob;
    } catch (e) {
      alert(`Falha ao preencher: ${e.message}`);
      return null;
    } finally {
      setLoading(false);
      setStatusMsg('');
    }
  };

  // Trigger browser download of DOCX
  const downloadDocx = async () => {
    let blob = generatedDocxBlob;
    if (!blob) {
      blob = await generateDocx();
    }
    if (!blob) return;

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Plano_de_Ensino_${planData.curso.replace(/\s+/g, '_')}.docx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  };

  // Convert DOCX to PDF on client side and download
  const downloadPdf = async () => {
    let blob = generatedDocxBlob;
    if (!blob) {
      blob = await generateDocx();
    }
    if (!blob) return;

    setLoading(true);
    setStatusMsg('Gerando o PDF final... Aguarde, isso pode levar alguns segundos.');

    try {
      // 1. Get the container div
      const previewContainer = document.getElementById('docx-preview-container');
      previewContainer.style.display = 'block'; // Make it temporarily block
      previewContainer.innerHTML = '<p style="text-align:center; padding: 20px;">Carregando documento para PDF...</p>';
      
      // 2. Render DOCX using docx-preview
      const arrayBuffer = await blob.arrayBuffer();
      await renderAsync(arrayBuffer, previewContainer, null, {
        inWrapper: false,
        ignoreWidth: false,
        ignoreHeight: false
      });

      // Give images/layouts a brief moment to render completely
      await new Promise(resolve => setTimeout(resolve, 800));

      // 3. Convert to PDF using html2pdf
      const opt = {
        margin: [0.4, 0.4, 0.4, 0.4], // 10mm margins for nice page fitting
        filename: `Plano_de_Ensino_${planData.curso.replace(/\s+/g, '_')}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, letterRendering: true },
        jsPDF: { unit: 'in', format: 'a4', orientation: 'portrait' },
        pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
      };

      if (!window.html2pdf) {
        throw new Error('A biblioteca de geração de PDF ainda não foi carregada no navegador. Aguarde alguns instantes.');
      }

      await window.html2pdf().from(previewContainer).set(opt).save();

    } catch (e) {
      alert(`Erro ao gerar PDF: ${e.message}`);
    } finally {
// 4. Hide the container again
      const previewContainer = document.getElementById('docx-preview-container');
      previewContainer.style.display = 'none';
      previewContainer.innerHTML = '';
      setLoading(false);
      setStatusMsg('');
    }
  };

  return (
    <div className="layout-container-three-column">
      {/* Background Orbs */}
      <div className="glowing-orb orb-purple"></div>
      <div className="glowing-orb orb-cyan"></div>

      {/* Hidden Div for Rendering DOCX to PDF */}
      <div id="docx-preview-container"></div>

      {/* COLUMN 1: Sources Manager (Left) */}
      <aside className="sources-panel">
        <div className="brand-header">
          <h1 style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '20px', fontWeight: '800' }}>
            <BrainCircuit style={{ color: 'hsl(var(--primary))' }} /> GusPlan
          </h1>
          <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))', marginTop: '2px' }}>
            Assistente MESEP SENAI
          </p>
        </div>

        <div className="sources-list-container">
          <h3 className="section-title">Fontes de Entrada</h3>
          
          {/* Source 1: Plano de Curso */}
          <div className="source-card">
            <div className="source-card-header">
              <span className={`status-dot ${coursePlanStatus}`}></span>
              <h4>1. Plano de Curso (PDF)</h4>
            </div>
            <p className="source-card-desc">Contém os objetivos, capacidades e ementa.</p>
            <label className="source-upload-btn">
              <input type="file" accept=".pdf" style={{ display: 'none' }} onChange={(e) => handleFileUpload(e, 'course_plan')} />
              {coursePlanStatus === 'ready' ? '✓ Alterar PDF' : coursePlanStatus === 'reading' ? 'Lendo...' : 'Carregar PDF'}
            </label>
            {coursePlanFile && <span className="source-filename">{coursePlanFile.name}</span>}
          </div>

          {/* Source 2: Guia MESEP */}
          <div className="source-card">
            <div className="source-card-header">
              <span className={`status-dot ${mesepBookStatus === 'ready' ? 'ready' : 'pending'}`}></span>
              <h4>2. Guia MESEP (PDF Opcional)</h4>
            </div>
            <p className="source-card-desc">Orientações metodológicas oficiais.</p>
            <label className="source-upload-btn">
              <input type="file" accept=".pdf" style={{ display: 'none' }} onChange={(e) => handleFileUpload(e, 'mesep_book')} />
              {mesepBookStatus === 'ready' ? '✓ Alterar PDF' : mesepBookStatus === 'reading' ? 'Lendo...' : 'Carregar PDF'}
            </label>
            {mesepBookFile && <span className="source-filename">{mesepBookFile.name}</span>}
          </div>

          {/* Source 3: Modelo DOCX */}
          <div className="source-card">
            <div className="source-card-header">
              <span className={`status-dot ${templateStatus}`}></span>
              <h4>3. Modelo DOCX</h4>
            </div>
            <p className="source-card-desc">Template oficial para preenchimento.</p>
            <label className="source-upload-btn">
              <input type="file" accept=".docx" style={{ display: 'none' }} onChange={(e) => handleFileUpload(e, 'template')} />
              {templateStatus === 'ready' ? '✓ Alterar Template' : 'Carregar DOCX'}
            </label>
            {templateFile && <span className="source-filename">{templateFile.name}</span>}
          </div>
        </div>

        {/* Generate Plan Button */}
        <div style={{ marginTop: 'auto', paddingTop: '16px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
          {statusMsg && (
            <div style={{ fontSize: '11px', color: 'hsl(var(--text-muted))', textAlign: 'center' }}>
              {statusMsg}
            </div>
          )}
          <button 
            className="btn-primary"
            disabled={coursePlanStatus !== 'ready' || templateStatus !== 'ready' || loading}
            onClick={analyzeCoursePlan}
            style={{ width: '100%', justifyContent: 'center' }}
          >
            {loading ? (
              <>
                <Loader2 className="animate-spin" size={16} /> Estruturando...
              </>
            ) : (
              <>
                <BrainCircuit size={16} /> Gerar Plano de Ensino
              </>
            )}
          </button>
        </div>
      </aside>

      {/* COLUMN 2: Chat Conversacional (Center) */}
      <section className="chat-panel">
        <div className="panel-header">
          <div>
            <h3 style={{ fontSize: '15px', fontWeight: '700' }}>Chat com as Fontes</h3>
            <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))' }}>Consulte e debata com base nos documentos enviados.</p>
          </div>
          <span className="badge">
            <BrainCircuit size={11} /> Gemini 2.0 Flash Lite
          </span>
        </div>

        <div className="chat-thread-container">
          {chatHistory.length === 0 ? (
            <div className="chat-placeholder">
              <p style={{ color: 'hsl(var(--text-muted))', fontSize: '13px', textAlign: 'center', marginBottom: '20px' }}>
                Olá! Faça perguntas sobre os documentos que você enviou. Sugestões:
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', width: '100%' }}>
                <button className="btn-secondary" style={{ textAlign: 'left', fontSize: '11px' }} onClick={() => handleSuggestionClick('Quais são as capacidades técnicas listadas no plano de curso?')}>
                  Quais as capacidades técnicas no plano de curso?
                </button>
                <button className="btn-secondary" style={{ textAlign: 'left', fontSize: '11px' }} onClick={() => handleSuggestionClick('Crie uma sugestão de desafio de Situação-Problema baseada nesse plano de curso.')}>
                  Crie um desafio de Situação-Problema.
                </button>
                <button className="btn-secondary" style={{ textAlign: 'left', fontSize: '11px' }} onClick={() => handleSuggestionClick('Explique como o MESEP aborda o desenvolvimento metodológico das situações de aprendizagem.')}>
                  Como o MESEP aborda as Situações de Aprendizagem?
                </button>
              </div>
            </div>
          ) : (
            <div className="chat-messages-scroll">
              {chatHistory.map((msg, idx) => (
                <div 
                  key={idx} 
                  style={{ 
                    alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start',
                    maxWidth: '85%',
                    background: msg.role === 'user' ? 'linear-gradient(135deg, hsl(var(--primary)) 0%, #6d28d9 100%)' : 'rgba(255, 255, 255, 0.05)',
                    border: msg.role === 'user' ? 'none' : '1px solid rgba(255,255,255,0.08)',
                    borderRadius: '12px',
                    padding: '10px 14px',
                    fontSize: '13px',
                    whiteSpace: 'pre-wrap',
                    boxShadow: msg.role === 'user' ? '0 4px 12px rgba(139,92,246,0.15)' : 'none',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '4px'
                  }}
                >
                  <span style={{ fontSize: '10px', fontWeight: '700', color: msg.role === 'user' ? '#e0d3ff' : '#a78bfa' }}>
                    {msg.role === 'user' ? 'Professor' : 'GusPlan IA'}
                  </span>
                  {msg.text}
                </div>
              ))}
              {chatLoading && (
                <div style={{ alignSelf: 'flex-start', background: 'rgba(255, 255, 255, 0.03)', border: '1px dashed rgba(255,255,255,0.1)', borderRadius: '12px', padding: '10px 14px', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <Loader2 className="animate-spin" size={14} /> Elaborando resposta...
                </div>
              )}
              <div ref={chatBottomRef} />
            </div>
          )}
        </div>

        <form onSubmit={sendChatMessage} className="chat-input-area">
          <input 
            type="text" 
            className="glass-input" 
            placeholder="Pergunte aos documentos..." 
            value={chatMessage} 
            onChange={(e) => setChatMessage(e.target.value)}
            disabled={chatLoading}
            style={{ flex: 1 }}
          />
          <button className="btn-primary" type="submit" disabled={chatLoading || !chatMessage.trim()}>
            <Send size={14} />
          </button>
        </form>
      </section>

      {/* COLUMN 3: Studio / Editor (Right) */}
      <section className="editor-panel">
        {!planData ? (
          <div className="editor-placeholder">
            <div className="placeholder-card">
              <BrainCircuit size={48} className="pulse-icon" style={{ color: 'hsl(var(--primary))', marginBottom: '16px' }} />
              <h3>Estúdio de Criação</h3>
              <p style={{ fontSize: '13px', color: 'hsl(var(--text-muted))', textAlign: 'center', margin: '8px 0 20px 0', lineHeight: '1.5' }}>
                Os planos estruturados de ensino aparecerão aqui. Envie seu Plano de Curso e o Modelo DOCX no menu esquerdo e clique em **Gerar Plano de Ensino** para começar.
              </p>
            </div>
          </div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
            {/* Editor Header */}
            <div className="panel-header" style={{ justifyContent: 'space-between' }}>
              <div>
                <h3 style={{ fontSize: '15px', fontWeight: '700' }}>Plano de Ensino Estruturado</h3>
                <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))' }}>Edite os dados sugeridos antes do preenchimento final.</p>
              </div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button className="btn-secondary" style={{ padding: '6px 12px', fontSize: '12px' }} onClick={downloadDocx} disabled={loading}>
                  <Download size={12} /> DOCX
                </button>
                <button className="btn-primary" style={{ padding: '6px 12px', fontSize: '12px' }} onClick={downloadPdf} disabled={loading}>
                  {loading ? <Loader2 className="animate-spin" size={12} /> : <Download size={12} />} Baixar PDF
                </button>
              </div>
            </div>

            {/* Editor Fields */}
            <div className="editor-fields-scroll">
              {/* Course Identity Details */}
              <div className="glass-panel" style={{ padding: '16px', display: 'grid', gridTemplateColumns: '2fr 1fr 1fr', gap: '12px', marginBottom: '16px' }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <label className="input-label">Curso</label>
                  <input type="text" className="glass-input" value={planData.curso} onChange={(e) => updatePlanField('curso', e.target.value)} />
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <label className="input-label">CH da UC</label>
                  <input type="text" className="glass-input" value={planData.carga_horaria_uc} onChange={(e) => updatePlanField('carga_horaria_uc', e.target.value)} />
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <label className="input-label">Nº de Aulas</label>
                  <input type="text" className="glass-input" value={planData.n_aulas} onChange={(e) => updatePlanField('n_aulas', e.target.value)} />
                </div>
                <div style={{ gridColumn: 'span 3', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <label className="input-label">Unidade Curricular (UC)</label>
                  <input type="text" className="glass-input" value={planData.unidade_curricular} onChange={(e) => updatePlanField('unidade_curricular', e.target.value)} />
                </div>
                <div style={{ gridColumn: 'span 3', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <label className="input-label">Objetivo da UC</label>
                  <textarea rows={2} className="glass-input" style={{ resize: 'vertical' }} value={planData.objetivo_uc} onChange={(e) => updatePlanField('objetivo_uc', e.target.value)} />
                </div>
              </div>

              {/* SA Navigation tabs */}
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <h4 style={{ fontSize: '13px', fontWeight: '700' }}>Situações de Aprendizagem</h4>
                <button className="btn-secondary" style={{ padding: '4px 8px', fontSize: '11px' }} onClick={addSa}>
                  <Plus size={11} /> Adicionar SA
                </button>
              </div>

              <div style={{ display: 'flex', gap: '6px', borderBottom: '1px solid rgba(255,255,255,0.06)', paddingBottom: '1px', marginBottom: '16px' }}>
                {planData.situacoes_aprendizagem.map((sa, idx) => (
                  <div key={idx} style={{ display: 'flex', alignItems: 'center' }}>
                    <button 
                      type="button"
                      className={`btn-secondary ${activeSaTab === idx ? 'active' : ''}`}
                      onClick={() => setActiveSaTab(idx)}
                      style={{ 
                        borderBottomLeftRadius: 0,
                        borderBottomRightRadius: 0,
                        background: activeSaTab === idx ? 'rgba(139, 92, 246, 0.1)' : 'transparent',
                        borderColor: activeSaTab === idx ? 'hsl(var(--primary)) hsl(var(--primary)) transparent' : 'transparent',
                        padding: '8px 12px',
                        fontSize: '12px'
                      }}
                    >
                      SA {idx + 1}
                    </button>
                    <button 
                      type="button"
                      style={{ background: 'transparent', border: 'none', color: '#ef4444', padding: '0 4px', cursor: 'pointer' }}
                      onClick={() => removeSa(idx)}
                    >
                      <Trash2 size={11} />
                    </button>
                  </div>
                ))}
              </div>

              {/* Edit Selected SA Fields */}
              {planData.situacoes_aprendizagem[activeSaTab] && (() => {
                const sa = planData.situacoes_aprendizagem[activeSaTab];
                return (
                  <div className="glass-panel" style={{ padding: '16px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '12px' }}>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                        <label className="input-label">Título da SA</label>
                        <input type="text" className="glass-input" value={sa.titulo} onChange={(e) => updateSaField(activeSaTab, 'titulo', e.target.value)} />
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                        <label className="input-label">CH Prevista</label>
                        <input type="text" className="glass-input" value={sa.carga_horaria} onChange={(e) => updateSaField(activeSaTab, 'carga_horaria', e.target.value)} />
                      </div>
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <label className="input-label">Estratégia de Aprendizagem</label>
                      <select 
                        className="glass-input" 
                        value={sa.estrategia_tipo} 
                        onChange={(e) => updateSaField(activeSaTab, 'estrategia_tipo', e.target.value)}
                        style={{ background: '#09090b', color: '#fff' }}
                      >
                        <option value="Situação-problema">Situação-problema</option>
                        <option value="Estudo de caso">Estudo de caso</option>
                        <option value="Pesquisa Aplicada">Pesquisa Aplicada</option>
                        <option value="Projeto">Projeto</option>
                        <option value="Integrador">Integrador</option>
                      </select>
                    </div>

                    {/* Technical Capacities */}
                    <div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                        <label className="input-label">Capacidades Técnicas</label>
                        <button type="button" className="btn-secondary" style={{ padding: '2px 6px', fontSize: '10px' }} onClick={() => addListItem(activeSaTab, 'capacidades_tecnicas', 'Nova capacidade técnica')}>
                          <Plus size={10} />
                        </button>
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        {sa.capacidades_tecnicas.map((cap, cIdx) => (
                          <div key={cIdx} style={{ display: 'flex', gap: '6px' }}>
                            <input type="text" className="glass-input" style={{ flex: 1, fontSize: '12px' }} value={cap} onChange={(e) => handleListChange(activeSaTab, 'capacidades_tecnicas', cIdx, e.target.value)} />
                            <button type="button" style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }} onClick={() => removeListItem(activeSaTab, 'capacidades_tecnicas', cIdx)}>
                              <Trash2 size={12} />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Socioemotional Capacities */}
                    <div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                        <label className="input-label">Capacidades Socioemocionais</label>
                        <button type="button" className="btn-secondary" style={{ padding: '2px 6px', fontSize: '10px' }} onClick={() => addListItem(activeSaTab, 'capacidades_socioemocionais', 'Nova capacidade socioemocional')}>
                          <Plus size={10} />
                        </button>
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        {sa.capacidades_socioemocionais.map((cap, cIdx) => (
                          <div key={cIdx} style={{ display: 'flex', gap: '6px' }}>
                            <input type="text" className="glass-input" style={{ flex: 1, fontSize: '12px' }} value={cap} onChange={(e) => handleListChange(activeSaTab, 'capacidades_socioemocionais', cIdx, e.target.value)} />
                            <button type="button" style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }} onClick={() => removeListItem(activeSaTab, 'capacidades_socioemocionais', cIdx)}>
                              <Trash2 size={12} />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Conhecimentos */}
                    <div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                        <label className="input-label">Conhecimentos</label>
                        <button type="button" className="btn-secondary" style={{ padding: '2px 6px', fontSize: '10px' }} onClick={() => addListItem(activeSaTab, 'conhecimentos', 'Novo conhecimento')}>
                          <Plus size={10} />
                        </button>
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        {sa.conhecimentos.map((kw, kIdx) => (
                          <div key={kIdx} style={{ display: 'flex', gap: '6px' }}>
                            <input type="text" className="glass-input" style={{ flex: 1, fontSize: '12px' }} value={kw} onChange={(e) => handleListChange(activeSaTab, 'conhecimentos', kIdx, e.target.value)} />
                            <button type="button" style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }} onClick={() => removeListItem(activeSaTab, 'conhecimentos', kIdx)}>
                              <Trash2 size={12} />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Contextualization, Desafio, Observações */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <label className="input-label">Contextualização</label>
                      <textarea rows={3} className="glass-input" style={{ fontSize: '12px', resize: 'vertical' }} value={sa.contextualizacao} onChange={(e) => updateSaField(activeSaTab, 'contextualizacao', e.target.value)} />
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <label className="input-label">Observações para o Docente</label>
                      <textarea rows={3} className="glass-input" style={{ fontSize: '12px', resize: 'vertical' }} value={sa.observacoes_docente} onChange={(e) => updateSaField(activeSaTab, 'observacoes_docente', e.target.value)} />
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <label className="input-label">Desafio da Situação de Aprendizagem</label>
                      <textarea rows={3} className="glass-input" style={{ fontSize: '12px', resize: 'vertical' }} value={sa.desafio} onChange={(e) => updateSaField(activeSaTab, 'desafio', e.target.value)} />
                    </div>

                    {/* Resultados Esperados */}
                    <div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                        <label className="input-label">Resultados Esperados</label>
                        <button type="button" className="btn-secondary" style={{ padding: '2px 6px', fontSize: '10px' }} onClick={() => addListItem(activeSaTab, 'resultados_esperados', 'Novo resultado esperado')}>
                          <Plus size={10} />
                        </button>
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        {sa.resultados_esperados.map((res, rIdx) => (
                          <div key={rIdx} style={{ display: 'flex', gap: '6px' }}>
                            <input type="text" className="glass-input" style={{ flex: 1, fontSize: '12px' }} value={res} onChange={(e) => handleListChange(activeSaTab, 'resultados_esperados', rIdx, e.target.value)} />
                            <button type="button" style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }} onClick={() => removeListItem(activeSaTab, 'resultados_esperados', rIdx)}>
                              <Trash2 size={12} />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Anexos & Referências */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <label className="input-label">Anexos (Figuras, leiautes, etc.)</label>
                      <input type="text" className="glass-input" value={sa.anexos} onChange={(e) => updateSaField(activeSaTab, 'anexos', e.target.value)} />
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <label className="input-label">Referências</label>
                      <input type="text" className="glass-input" value={sa.referencias} onChange={(e) => updateSaField(activeSaTab, 'referencias', e.target.value)} />
                    </div>

                    {/* Evaluation Criteria Instrument Table */}
                    <div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                        <label className="input-label">Instrumento de Avaliação (Capacidades e Critérios)</label>
                        <button type="button" className="btn-primary" style={{ padding: '4px 8px', fontSize: '10px' }} onClick={() => addInstrumentRow(activeSaTab)}>
                          <Plus size={10} /> Adicionar Critério
                        </button>
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        {sa.instrumento_registro.map((row, iIdx) => (
                          <div key={iIdx} className="glass-panel" style={{ padding: '10px', display: 'flex', flexDirection: 'column', gap: '8px', border: '1px solid rgba(255,255,255,0.04)' }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                              <span style={{ fontSize: '10px', color: 'hsl(var(--primary))', fontWeight: '700' }}>Critério {iIdx + 1}</span>
                              <button type="button" style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }} onClick={() => removeInstrumentRow(activeSaTab, iIdx)}>
                                <Trash2 size={11} />
                              </button>
                            </div>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                              <label style={{ fontSize: '10px', color: 'hsl(var(--text-muted))' }}>Capacidade Associada</label>
                              <input type="text" className="glass-input" style={{ fontSize: '11px', padding: '6px 10px' }} value={row.capacidade} onChange={(e) => handleInstrumentChange(activeSaTab, iIdx, 'capacidade', e.target.value)} />
                            </div>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                              <label style={{ fontSize: '10px', color: 'hsl(var(--text-muted))' }}>Pergunta do Critério (Comece com [CRÍTICO] para destacar)</label>
                              <input type="text" className="glass-input" style={{ fontSize: '11px', padding: '6px 10px' }} value={row.criterio} onChange={(e) => handleInstrumentChange(activeSaTab, iIdx, 'criterio', e.target.value)} />
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
          </div>
        )}
      </section>

      {/* Floating Status / Loader Modal */}
      {loading && (
        <div style={{ position: 'fixed', bottom: '24px', right: '24px', background: 'rgba(9,9,11,0.9)', border: '1px solid rgba(255,255,255,0.08)', padding: '12px 20px', borderRadius: '12px', zIndex: 100, backdropFilter: 'blur(10px)', boxShadow: '0 10px 30px rgba(0,0,0,0.5)', display: 'flex', flexDirection: 'column', gap: '4px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', fontWeight: 'bold' }}>
            <Loader2 className="animate-spin" size={16} style={{ color: 'hsl(var(--primary))' }} />
            Processando requisição...
          </div>
          <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))' }}>{statusMsg}</p>
        </div>
      )}
    </div>
  );
}
