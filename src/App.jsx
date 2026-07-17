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
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('gemini_api_key') || '');
  const [modelName, setModelName] = useState('gemini-2.0-flash-lite');
  const [customModel, setCustomModel] = useState('');
  const [activeTab, setActiveTab] = useState('settings'); // settings, sources, chat, editor
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

  // Save API key
  useEffect(() => {
    localStorage.setItem('gemini_api_key', apiKey);
  }, [apiKey]);

  // Scroll chat to bottom
  useEffect(() => {
    chatBottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [chatHistory]);

  const getActiveModel = () => {
    return modelName === 'custom' ? customModel : modelName;
  };

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
    if (!apiKey) {
      alert('Por favor, configure sua API Key do Gemini na aba Configurações.');
      setActiveTab('settings');
      return;
    }

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
      formData.append('api_key', apiKey);
      formData.append('model_name', getActiveModel());

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
      setActiveTab('sources');
      return;
    }
    if (!apiKey) {
      alert('Por favor, insira sua API Key do Gemini na aba Configurações.');
      setActiveTab('settings');
      return;
    }

    setLoading(true);
    setStatusMsg('Analisando o Plano de Curso com IA...');
    try {
      const formData = new FormData();
      formData.append('course_plan', coursePlanText);
      formData.append('mesep_book', mesepBookText);
      formData.append('api_key', apiKey);
      formData.append('model_name', getActiveModel());

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
      setActiveTab('editor');
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
      setActiveTab('sources');
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
    <div className="layout-container">
      {/* Background Orbs */}
      <div className="glowing-orb orb-purple"></div>
      <div className="glowing-orb orb-cyan"></div>

      {/* Hidden Div for Rendering DOCX to PDF */}
      <div id="docx-preview-container"></div>

      {/* LEFT PANEL: Sidebar & Sources */}
      <aside className="sidebar">
        <div>
          <h1 style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '24px', fontWeight: '800' }}>
            <BrainCircuit style={{ color: 'hsl(var(--primary))' }} /> GusPlan
          </h1>
          <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))', marginTop: '4px' }}>
            Assistente MESEP SENAI
          </p>
        </div>

        {/* Navigation Tabs */}
        <nav style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
          <button 
            className={`btn-secondary ${activeTab === 'settings' ? 'active' : ''}`}
            onClick={() => setActiveTab('settings')}
            style={{ 
              justifyContent: 'flex-start',
              background: activeTab === 'settings' ? 'rgba(139, 92, 246, 0.15)' : 'transparent',
              borderColor: activeTab === 'settings' ? 'rgba(139, 92, 246, 0.3)' : 'transparent'
            }}
          >
            <Settings size={18} style={{ color: activeTab === 'settings' ? 'hsl(var(--primary))' : 'inherit' }} />
            Configurações
          </button>
          
          <button 
            className={`btn-secondary ${activeTab === 'sources' ? 'active' : ''}`}
            onClick={() => setActiveTab('sources')}
            style={{ 
              justifyContent: 'flex-start',
              background: activeTab === 'sources' ? 'rgba(139, 92, 246, 0.15)' : 'transparent',
              borderColor: activeTab === 'sources' ? 'rgba(139, 92, 246, 0.3)' : 'transparent'
            }}
          >
            <FolderOpen size={18} style={{ color: activeTab === 'sources' ? 'hsl(var(--primary))' : 'inherit' }} />
            Arquivos Fontes
          </button>

          <button 
            className={`btn-secondary ${activeTab === 'chat' ? 'active' : ''}`}
            onClick={() => setActiveTab('chat')}
            style={{ 
              justifyContent: 'flex-start',
              background: activeTab === 'chat' ? 'rgba(139, 92, 246, 0.15)' : 'transparent',
              borderColor: activeTab === 'chat' ? 'rgba(139, 92, 246, 0.3)' : 'transparent'
            }}
          >
            <Send size={18} style={{ color: activeTab === 'chat' ? 'hsl(var(--primary))' : 'inherit' }} />
            Chat Conversacional
          </button>

          <button 
            className={`btn-secondary ${activeTab === 'editor' ? 'active' : ''}`}
            disabled={!planData}
            onClick={() => setActiveTab('editor')}
            style={{ 
              justifyContent: 'flex-start',
              background: activeTab === 'editor' ? 'rgba(139, 92, 246, 0.15)' : 'transparent',
              borderColor: activeTab === 'editor' ? 'rgba(139, 92, 246, 0.3)' : 'transparent'
            }}
          >
            <FileCheck size={18} style={{ color: activeTab === 'editor' ? 'hsl(var(--primary))' : 'inherit' }} />
            Editor do Plano
            {planData && <span className="badge">Pronto</span>}
          </button>
        </nav>

        {/* File status overview */}
        <div className="glass-panel" style={{ marginTop: 'auto', display: 'flex', flexDirection: 'column', gap: '10px', padding: '16px', borderRadius: '12px' }}>
          <h4 style={{ fontSize: '13px', marginBottom: '8px' }}>Status das Fontes</h4>
          
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px' }}>
            <div style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: coursePlanStatus === 'ready' ? '#10b981' : coursePlanStatus === 'reading' ? '#06b6d4' : '#ef4444' }}></div>
            <span style={{ color: coursePlanStatus === 'ready' ? '#fff' : 'hsl(var(--text-muted))' }}>Plano de Curso</span>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px' }}>
            <div style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: mesepBookStatus === 'ready' ? '#10b981' : mesepBookStatus === 'reading' ? '#06b6d4' : 'rgba(255,255,255,0.2)' }}></div>
            <span style={{ color: mesepBookStatus === 'ready' ? '#fff' : 'hsl(var(--text-muted))' }}>Livro MESEP (Opcional)</span>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px' }}>
            <div style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: templateStatus === 'ready' ? '#10b981' : '#ef4444' }}></div>
            <span style={{ color: templateStatus === 'ready' ? '#fff' : 'hsl(var(--text-muted))' }}>Modelo DOCX</span>
          </div>
        </div>
      </aside>

      {/* CENTRAL AREA: Chat Console or Settings or Upload Panel */}
      <main className="main-content">
        
        {/* Active Tab rendering */}
        {activeTab === 'settings' && (
          <div style={{ padding: '40px', display: 'flex', flexDirection: 'column', gap: '24px', overflowY: 'auto' }}>
            <div>
              <h2 style={{ fontSize: '28px', marginBottom: '8px' }}>Configurações da Inteligência Artificial</h2>
              <p style={{ color: 'hsl(var(--text-muted))', fontSize: '14px' }}>Configure a API do Gemini para gerar seus planos de ensino sob medida.</p>
            </div>

            <div className="glass-panel" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <label style={{ fontSize: '13px', fontWeight: '600' }}>Chave da API do Gemini</label>
                <input 
                  type="password"
                  className="glass-input"
                  placeholder="Insira sua Gemini API Key (AIzaSy...)"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                />
                <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))' }}>Sua chave é armazenada apenas no seu navegador local.</p>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <label style={{ fontSize: '13px', fontWeight: '600' }}>Modelo do Gemini</label>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                  <button 
                    className={`btn-secondary ${modelName === 'gemini-2.0-flash-lite' ? 'active' : ''}`}
                    onClick={() => setModelName('gemini-2.0-flash-lite')}
                    style={{ background: modelName === 'gemini-2.0-flash-lite' ? 'rgba(139, 92, 246, 0.15)' : 'transparent', borderColor: modelName === 'gemini-2.0-flash-lite' ? 'hsl(var(--primary))' : 'rgba(255,255,255,0.1)' }}
                  >
                    Gemini 2.0 Flash Lite (Recomendado)
                  </button>
                  <button 
                    className={`btn-secondary ${modelName === 'custom' ? 'active' : ''}`}
                    onClick={() => setModelName('custom')}
                    style={{ background: modelName === 'custom' ? 'rgba(139, 92, 246, 0.15)' : 'transparent', borderColor: modelName === 'custom' ? 'hsl(var(--primary))' : 'rgba(255,255,255,0.1)' }}
                  >
                    Outro (Personalizado)
                  </button>
                </div>

                {modelName === 'custom' && (
                  <input 
                    type="text"
                    className="glass-input"
                    style={{ marginTop: '8px' }}
                    placeholder="Digite o nome do modelo (ex: gemini-1.5-flash-lite-preview-0514)"
                    value={customModel}
                    onChange={(e) => setCustomModel(e.target.value)}
                  />
                )}
              </div>
            </div>

            <button 
              className="btn-primary" 
              style={{ width: 'fit-content' }}
              onClick={() => setActiveTab('sources')}
            >
              Próximo: Carregar Arquivos <ChevronRight size={16} />
            </button>
          </div>
        )}

        {activeTab === 'sources' && (
          <div style={{ padding: '40px', display: 'flex', flexDirection: 'column', gap: '24px', overflowY: 'auto' }}>
            <div>
              <h2 style={{ fontSize: '28px', marginBottom: '8px' }}>Carregar Arquivos Fontes</h2>
              <p style={{ color: 'hsl(var(--text-muted))', fontSize: '14px' }}>Envie os documentos oficiais para alimentar a IA e preencher o template.</p>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '20px' }}>
              {/* File 1: Plano de Curso */}
              <div className="glass-panel" style={{ padding: '20px' }}>
                <h3 style={{ fontSize: '16px', marginBottom: '4px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <FileText style={{ color: 'hsl(var(--secondary))' }} /> 1. Plano de Curso (PDF do Curso a ser convertido)
                </h3>
                <p style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', marginBottom: '16px' }}>
                  Contém os objetivos, as capacidades técnicas e socioemocionais e a ementa de conteúdos.
                </p>
                <div className={`upload-box ${coursePlanStatus === 'ready' ? 'uploaded' : ''}`}>
                  <label style={{ display: 'block', cursor: 'pointer' }}>
                    <input 
                      type="file" 
                      accept=".pdf" 
                      style={{ display: 'none' }} 
                      onChange={(e) => handleFileUpload(e, 'course_plan')}
                    />
                    {coursePlanStatus === 'ready' ? (
                      <div>
                        <p style={{ color: '#10b981', fontWeight: '600' }}>✓ Plano de Curso Carregado com Sucesso</p>
                        <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))', marginTop: '4px' }}>{coursePlanFile?.name}</p>
                      </div>
                    ) : coursePlanStatus === 'reading' ? (
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}>
                        <Loader2 className="animate-spin" size={18} />
                        <p>Extraindo conteúdo do PDF...</p>
                      </div>
                    ) : (
                      <p>Arraste ou clique para enviar o PDF do Plano de Curso</p>
                    )}
                  </label>
                </div>
              </div>

              {/* File 2: Book MESEP */}
              <div className="glass-panel" style={{ padding: '20px' }}>
                <h3 style={{ fontSize: '16px', marginBottom: '4px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <FileText style={{ color: 'hsl(var(--primary))' }} /> 2. Guia/Livro MESEP (Opcional - PDF)
                </h3>
                <p style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', marginBottom: '16px' }}>
                  Livro de diretrizes da metodologia. Se não enviado, a IA utilizará as regras integradas em seu prompt de sistema.
                </p>
                <div className={`upload-box ${mesepBookStatus === 'ready' ? 'uploaded' : ''}`}>
                  <label style={{ display: 'block', cursor: 'pointer' }}>
                    <input 
                      type="file" 
                      accept=".pdf" 
                      style={{ display: 'none' }} 
                      onChange={(e) => handleFileUpload(e, 'mesep_book')}
                    />
                    {mesepBookStatus === 'ready' ? (
                      <div>
                        <p style={{ color: '#10b981', fontWeight: '600' }}>✓ Diretrizes MESEP Carregadas</p>
                        <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))', marginTop: '4px' }}>{mesepBookFile?.name}</p>
                      </div>
                    ) : mesepBookStatus === 'reading' ? (
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}>
                        <Loader2 className="animate-spin" size={18} />
                        <p>Extraindo conteúdo do Livro...</p>
                      </div>
                    ) : (
                      <p>Arraste ou clique para enviar o PDF do Livro MESEP (Opcional)</p>
                    )}
                  </label>
                </div>
              </div>

              {/* File 3: DOCX Template */}
              <div className="glass-panel" style={{ padding: '20px' }}>
                <h3 style={{ fontSize: '16px', marginBottom: '4px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <FileCheck style={{ color: 'hsl(var(--accent))' }} /> 3. Modelo do Arquivo (Template DOCX Oficial)
                </h3>
                <p style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', marginBottom: '16px' }}>
                  Arquivo .docx oficial do SENAI contendo a estrutura de tabelas vazias do plano de ensino.
                </p>
                <div className={`upload-box ${templateStatus === 'ready' ? 'uploaded' : ''}`}>
                  <label style={{ display: 'block', cursor: 'pointer' }}>
                    <input 
                      type="file" 
                      accept=".docx" 
                      style={{ display: 'none' }} 
                      onChange={(e) => handleFileUpload(e, 'template')}
                    />
                    {templateStatus === 'ready' ? (
                      <div>
                        <p style={{ color: '#10b981', fontWeight: '600' }}>✓ Template DOCX Selecionado</p>
                        <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))', marginTop: '4px' }}>{templateFile?.name}</p>
                      </div>
                    ) : (
                      <p>Arraste ou clique para enviar o arquivo Modelo DOCX</p>
                    )}
                  </label>
                </div>
              </div>
            </div>

            <div style={{ display: 'flex', gap: '16px', marginTop: '12px' }}>
              <button 
                className="btn-primary"
                disabled={coursePlanStatus !== 'ready' || templateStatus !== 'ready' || loading}
                onClick={analyzeCoursePlan}
              >
                {loading ? (
                  <>
                    <Loader2 className="animate-spin" size={16} /> Analisando e Gerando Rascunho...
                  </>
                ) : (
                  <>
                    <BrainCircuit size={16} /> Analisar e Estruturar com IA
                  </>
                )}
              </button>
              
              <button 
                className="btn-secondary"
                disabled={coursePlanStatus !== 'ready'}
                onClick={() => setActiveTab('chat')}
              >
                Conversar com Fontes (Chat) <ChevronRight size={16} />
              </button>
            </div>
          </div>
        )}

        {activeTab === 'chat' && (
          <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
            {/* Chat Header */}
            <div style={{ padding: '20px 24px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <h3 style={{ fontSize: '18px' }}>Chat Inteligente (NotebookLM Style)</h3>
                <p style={{ fontSize: '12px', color: 'hsl(var(--text-muted))' }}>Consulte, debata e faça perguntas com base no contexto dos arquivos que carregou.</p>
              </div>
              <span className="badge">
                <BrainCircuit size={12} /> {getActiveModel()}
              </span>
            </div>

            {/* Suggestions panel when history is empty */}
            {chatHistory.length === 0 && (
              <div style={{ padding: '40px 24px', display: 'flex', flexDirection: 'column', gap: '16px', overflowY: 'auto' }}>
                <p style={{ textAlign: 'center', color: 'hsl(var(--text-muted))', fontSize: '14px' }}>Sem histórico. Pergunte algo ou clique em uma das sugestões abaixo:</p>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', maxWidth: '800px', margin: '0 auto' }}>
                  <button className="btn-secondary" style={{ fontSize: '12px', textAlign: 'left' }} onClick={() => handleSuggestionClick('Quais são as capacidades técnicas listadas no plano de curso?')}>
                    Quais são as capacidades técnicas listadas no plano de curso?
                  </button>
                  <button className="btn-secondary" style={{ fontSize: '12px', textAlign: 'left' }} onClick={() => handleSuggestionClick('Crie uma sugestão de desafio de Situação-Problema baseada nesse plano de curso.')}>
                    Crie uma sugestão de desafio de Situação-Problema baseada nesse plano de curso.
                  </button>
                  <button className="btn-secondary" style={{ fontSize: '12px', textAlign: 'left' }} onClick={() => handleSuggestionClick('Como posso formular critérios de avaliação claros para a capacidade socioemocional de autonomia?')}>
                    Como formular critérios para a capacidade socioemocional de autonomia?
                  </button>
                  <button className="btn-secondary" style={{ fontSize: '12px', textAlign: 'left' }} onClick={() => handleSuggestionClick('Explique como o MESEP aborda o desenvolvimento metodológico das situações de aprendizagem.')}>
                    Explique como o MESEP aborda as situações de aprendizagem.
                  </button>
                </div>
              </div>
            )}

            {/* Chat Thread */}
            {chatHistory.length > 0 && (
              <div style={{ flex: 1, overflowY: 'auto', padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                {chatHistory.map((msg, idx) => (
                  <div 
                    key={idx} 
                    style={{ 
                      alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start',
                      maxWidth: '75%',
                      background: msg.role === 'user' ? 'linear-gradient(135deg, hsl(var(--primary)) 0%, #6d28d9 100%)' : 'rgba(255, 255, 255, 0.05)',
                      border: msg.role === 'user' ? 'none' : '1px solid rgba(255,255,255,0.08)',
                      borderRadius: '12px',
                      padding: '12px 16px',
                      fontSize: '14px',
                      whiteSpace: 'pre-wrap',
                      boxShadow: msg.role === 'user' ? '0 4px 12px rgba(139,92,246,0.2)' : 'none'
                    }}
                  >
                    <p style={{ fontWeight: '600', fontSize: '11px', color: msg.role === 'user' ? '#ddd' : '#a78bfa', marginBottom: '4px' }}>
                      {msg.role === 'user' ? 'Professor' : 'GusPlan IA'}
                    </p>
                    {msg.text}
                  </div>
                ))}
                
                {chatLoading && (
                  <div style={{ alignSelf: 'flex-start', background: 'rgba(255, 255, 255, 0.03)', border: '1px dashed rgba(255,255,255,0.1)', borderRadius: '12px', padding: '12px 16px', fontSize: '14px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Loader2 className="animate-spin" size={16} /> Pensando na resposta...
                  </div>
                )}
                <div ref={chatBottomRef} />
              </div>
            )}

            {/* Chat Input form */}
            <form onSubmit={sendChatMessage} style={{ padding: '16px 24px', borderTop: '1px solid rgba(255,255,255,0.06)', display: 'flex', gap: '12px' }}>
              <input 
                type="text" 
                className="glass-input" 
                style={{ flex: 1 }} 
                placeholder="Pergunte aos documentos oficiais..."
                value={chatMessage}
                onChange={(e) => setChatMessage(e.target.value)}
                disabled={chatLoading}
              />
              <button className="btn-primary" type="submit" disabled={chatLoading || !chatMessage.trim()}>
                <Send size={16} />
              </button>
            </form>
          </div>
        )}

        {activeTab === 'editor' && planData && (
          <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
            
            {/* Editor Header */}
            <div style={{ padding: '20px 24px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <h3 style={{ fontSize: '18px' }}>Editor de Rascunho do Plano de Ensino</h3>
                <p style={{ fontSize: '12px', color: 'hsl(var(--text-muted))' }}>Revise, adicione, remova ou edite qualquer dado gerado pela IA antes do preenchimento final do DOCX.</p>
              </div>
              
              <div style={{ display: 'flex', gap: '12px' }}>
                <button className="btn-secondary" onClick={downloadDocx} disabled={loading}>
                  <Download size={14} /> DOCX
                </button>
                <button className="btn-primary" onClick={downloadPdf} disabled={loading}>
                  {loading ? <Loader2 className="animate-spin" size={14} /> : <Download size={14} />} Baixar PDF Pronto
                </button>
              </div>
            </div>

            {/* Main Editor Tabs and Fields */}
            <div style={{ flex: 1, overflowY: 'auto', padding: '24px', display: 'flex', flexDirection: 'column', gap: '24px' }}>
              
              {/* Course Identity Details */}
              <div className="glass-panel" style={{ padding: '20px', display: 'grid', gridTemplateColumns: '2fr 1fr 1fr', gap: '16px' }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Curso</label>
                  <input type="text" className="glass-input" value={planData.curso} onChange={(e) => updatePlanField('curso', e.target.value)} />
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Carga Horária da UC</label>
                  <input type="text" className="glass-input" value={planData.carga_horaria_uc} onChange={(e) => updatePlanField('carga_horaria_uc', e.target.value)} />
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Nº de Aulas</label>
                  <input type="text" className="glass-input" value={planData.n_aulas} onChange={(e) => updatePlanField('n_aulas', e.target.value)} />
                </div>
                <div style={{ gridColumn: 'span 3', display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Unidade Curricular (UC)</label>
                  <input type="text" className="glass-input" value={planData.unidade_curricular} onChange={(e) => updatePlanField('unidade_curricular', e.target.value)} />
                </div>
                <div style={{ gridColumn: 'span 3', display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Objetivo da Unidade Curricular</label>
                  <textarea rows={3} className="glass-input" style={{ resize: 'vertical' }} value={planData.objetivo_uc} onChange={(e) => updatePlanField('objetivo_uc', e.target.value)} />
                </div>
              </div>

              {/* SA Navigation Tabs */}
              <div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                  <h4 style={{ fontSize: '15px' }}>Situações de Aprendizagem</h4>
                  <button className="btn-secondary" style={{ padding: '6px 12px', fontSize: '12px' }} onClick={addSa}>
                    <Plus size={14} /> Adicionar SA
                  </button>
                </div>
                
                <div style={{ display: 'flex', gap: '8px', borderBottom: '1px solid rgba(255,255,255,0.06)', paddingBottom: '1px' }}>
                  {planData.situacoes_aprendizagem.map((sa, idx) => (
                    <div key={idx} style={{ display: 'flex', alignItems: 'center' }}>
                      <button 
                        className={`btn-secondary ${activeSaTab === idx ? 'active' : ''}`}
                        onClick={() => setActiveSaTab(idx)}
                        style={{ 
                          borderBottomLeftRadius: 0,
                          borderBottomRightRadius: 0,
                          background: activeSaTab === idx ? 'rgba(139, 92, 246, 0.1)' : 'transparent',
                          borderColor: activeSaTab === idx ? 'hsl(var(--primary)) hsl(var(--primary)) transparent' : 'transparent',
                          padding: '10px 16px',
                          fontSize: '13px'
                        }}
                      >
                        SA {idx + 1}
                      </button>
                      <button 
                        style={{ background: 'transparent', border: 'none', color: '#ef4444', padding: '0 8px', cursor: 'pointer' }}
                        onClick={() => removeSa(idx)}
                        title="Remover SA"
                      >
                        <Trash2 size={13} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              {/* Edit Selected SA Fields */}
              {planData.situacoes_aprendizagem[activeSaTab] && (() => {
                const sa = planData.situacoes_aprendizagem[activeSaTab];
                return (
                  <div className="glass-panel" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                    <div style={{ display: 'grid', gridTemplateColumns: '3fr 1fr 1fr', gap: '16px' }}>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Título da SA</label>
                        <input type="text" className="glass-input" value={sa.titulo} onChange={(e) => updateSaField(activeSaTab, 'titulo', e.target.value)} />
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Carga Horária Prevista</label>
                        <input type="text" className="glass-input" value={sa.carga_horaria} onChange={(e) => updateSaField(activeSaTab, 'carga_horaria', e.target.value)} />
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Estratégia Desafiadora</label>
                        <select 
                          className="glass-input"
                          style={{ background: '#111522' }}
                          value={sa.estrategia_tipo} 
                          onChange={(e) => updateSaField(activeSaTab, 'estrategia_tipo', e.target.value)}
                        >
                          <option value="Situação-problema">Situação-problema</option>
                          <option value="Estudo de caso">Estudo de caso</option>
                          <option value="Pesquisa Aplicada">Pesquisa Aplicada</option>
                          <option value="Projeto">Projeto</option>
                          <option value="Integrador">Integrador</option>
                        </select>
                      </div>
                    </div>

                    {/* Lists (Capacidades, Conhecimentos) */}
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                      {/* Capacidades Técnicas */}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Capacidades Técnicas</label>
                          <button className="btn-secondary" style={{ padding: '2px 8px', fontSize: '10px' }} onClick={() => addListItem(activeSaTab, 'capacidades_tecnicas', 'Nova capacidade')}>+ Add</button>
                        </div>
                        {sa.capacidades_tecnicas.map((item, lIdx) => (
                          <div key={lIdx} style={{ display: 'flex', gap: '8px' }}>
                            <input type="text" className="glass-input" style={{ flex: 1, padding: '8px 12px' }} value={item} onChange={(e) => handleListChange(activeSaTab, 'capacidades_tecnicas', lIdx, e.target.value)} />
                            <button style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }} onClick={() => removeListItem(activeSaTab, 'capacidades_tecnicas', lIdx)}><Trash2 size={14} /></button>
                          </div>
                        ))}
                      </div>

                      {/* Capacidades Socioemocionais */}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Capacidades Socioemocionais</label>
                          <button className="btn-secondary" style={{ padding: '2px 8px', fontSize: '10px' }} onClick={() => addListItem(activeSaTab, 'capacidades_socioemocionais', 'Nova capacidade socioemocional')}>+ Add</button>
                        </div>
                        {sa.capacidades_socioemocionais.map((item, lIdx) => (
                          <div key={lIdx} style={{ display: 'flex', gap: '8px' }}>
                            <input type="text" className="glass-input" style={{ flex: 1, padding: '8px 12px' }} value={item} onChange={(e) => handleListChange(activeSaTab, 'capacidades_socioemocionais', lIdx, e.target.value)} />
                            <button style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }} onClick={() => removeListItem(activeSaTab, 'capacidades_socioemocionais', lIdx)}><Trash2 size={14} /></button>
                          </div>
                        ))}
                      </div>

                      {/* Conhecimentos */}
                      <div style={{ gridColumn: 'span 2', display: 'flex', flexDirection: 'column', gap: '10px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Conhecimentos Relacionados</label>
                          <button className="btn-secondary" style={{ padding: '2px 8px', fontSize: '10px' }} onClick={() => addListItem(activeSaTab, 'conhecimentos', 'Novo conhecimento')}>+ Add</button>
                        </div>
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
                          {sa.conhecimentos.map((item, lIdx) => (
                            <div key={lIdx} style={{ display: 'flex', gap: '8px' }}>
                              <input type="text" className="glass-input" style={{ flex: 1, padding: '8px 12px' }} value={item} onChange={(e) => handleListChange(activeSaTab, 'conhecimentos', lIdx, e.target.value)} />
                              <button style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }} onClick={() => removeListItem(activeSaTab, 'conhecimentos', lIdx)}><Trash2 size={14} /></button>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>

                    {/* Pedagogy details (Context, Challenge, etc) */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Contextualização</label>
                        <textarea rows={4} className="glass-input" style={{ resize: 'vertical' }} value={sa.contextualizacao} onChange={(e) => updateSaField(activeSaTab, 'contextualizacao', e.target.value)} />
                      </div>

                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Observações para o Docente</label>
                        <textarea rows={3} className="glass-input" style={{ resize: 'vertical' }} value={sa.observacoes_docente} onChange={(e) => updateSaField(activeSaTab, 'observacoes_docente', e.target.value)} />
                      </div>

                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Desafio do Estudante</label>
                        <textarea rows={4} className="glass-input" style={{ resize: 'vertical' }} value={sa.desafio} onChange={(e) => updateSaField(activeSaTab, 'desafio', e.target.value)} />
                      </div>

                      {/* Resultados Esperados */}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Resultados Esperados (Entregáveis)</label>
                          <button className="btn-secondary" style={{ padding: '2px 8px', fontSize: '10px' }} onClick={() => addListItem(activeSaTab, 'resultados_esperados', 'Novo resultado esperado')}>+ Add</button>
                        </div>
                        {sa.resultados_esperados.map((item, lIdx) => (
                          <div key={lIdx} style={{ display: 'flex', gap: '8px' }}>
                            <input type="text" className="glass-input" style={{ flex: 1, padding: '8px 12px' }} value={item} onChange={(e) => handleListChange(activeSaTab, 'resultados_esperados', lIdx, e.target.value)} />
                            <button style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }} onClick={() => removeListItem(activeSaTab, 'resultados_esperados', lIdx)}><Trash2 size={14} /></button>
                          </div>
                        ))}
                      </div>

                      {/* Anexos e Referências */}
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Anexos sugeridos</label>
                          <input type="text" className="glass-input" value={sa.anexos} onChange={(e) => updateSaField(activeSaTab, 'anexos', e.target.value)} />
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <label style={{ fontSize: '12px', color: 'hsl(var(--text-muted))', fontWeight: '600' }}>Referências bibliográficas</label>
                          <input type="text" className="glass-input" value={sa.referencias} onChange={(e) => updateSaField(activeSaTab, 'referencias', e.target.value)} />
                        </div>
                      </div>
                    </div>

                    {/* Instrumento de Registro (Assessment Criteria) */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: '1px solid rgba(255,255,255,0.06)', paddingTop: '16px' }}>
                        <h4 style={{ fontSize: '14px', fontWeight: '600' }}>Instrumento de Registro & Critérios de Avaliação</h4>
                        <button className="btn-secondary" style={{ padding: '4px 10px', fontSize: '11px' }} onClick={() => addInstrumentRow(activeSaTab)}>+ Adicionar Critério</button>
                      </div>
                      
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        {sa.instrumento_registro.map((pair, instIdx) => (
                          <div key={instIdx} style={{ display: 'grid', gridTemplateColumns: '2fr 3fr auto', gap: '12px', alignItems: 'center' }}>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                              <label style={{ fontSize: '10px', color: 'hsl(var(--text-muted))' }}>Capacidade Relacionada</label>
                              <input type="text" className="glass-input" style={{ padding: '6px 10px', fontSize: '12px' }} value={pair.capacidade} onChange={(e) => handleInstrumentChange(activeSaTab, instIdx, 'capacidade', e.target.value)} />
                            </div>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                              <label style={{ display: 'flex', justifyContent: 'space-between', fontSize: '10px', color: 'hsl(var(--text-muted))' }}>
                                <span>Critério de Avaliação (Pergunta)</span>
                                <span style={{ color: pair.criterio.includes('[CRÍTICO]') ? '#ef4444' : '#a78bfa', fontWeight: 'bold' }}>
                                  {pair.criterio.includes('[CRÍTICO]') ? 'CRÍTICO (Negrito)' : 'Desejável'}
                                </span>
                              </label>
                              <input type="text" className="glass-input" style={{ padding: '6px 10px', fontSize: '12px' }} value={pair.criterio} onChange={(e) => handleInstrumentChange(activeSaTab, instIdx, 'criterio', e.target.value)} placeholder="Use [CRÍTICO] no início para critérios críticos" />
                            </div>
                            <button 
                              style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer', marginTop: '16px' }} 
                              onClick={() => removeInstrumentRow(activeSaTab, instIdx)}
                            >
                              <Trash2 size={16} />
                            </button>
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

      </main>

      {/* RIGHT PANEL: Help console & Document generation */}
      <aside className="right-panel">
        <div>
          <h3 style={{ fontSize: '16px', marginBottom: '8px' }}>Resumo & Ações</h3>
          <p style={{ fontSize: '12px', color: 'hsl(var(--text-muted))' }}>Carregue as fontes, confira o chat e edite para exportar.</p>
        </div>

        {/* Global actions */}
        <div className="glass-panel" style={{ padding: '20px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
          <h4 style={{ fontSize: '13px' }}>Ações Rápidas</h4>
          
          <button 
            className="btn-primary" 
            style={{ width: '100%', justifyContent: 'center' }}
            disabled={!planData || loading}
            onClick={downloadPdf}
          >
            {loading && statusMsg.includes('PDF') ? (
              <>
                <Loader2 className="animate-spin" size={16} /> Gerando PDF...
              </>
            ) : (
              <>
                <Download size={16} /> Baixar PDF Pronto (1:1)
              </>
            )}
          </button>

          <button 
            className="btn-secondary" 
            style={{ width: '100%', justifyContent: 'center' }}
            disabled={!planData || loading}
            onClick={downloadDocx}
          >
            {loading && statusMsg.includes('DOCX') ? (
              <>
                <Loader2 className="animate-spin" size={16} /> Preenchendo...
              </>
            ) : (
              <>
                <Download size={16} /> Baixar em Word (DOCX)
              </>
            )}
          </button>
        </div>

        {/* Loading modal/toast for generation */}
        {loading && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', padding: '16px', background: 'rgba(139, 92, 246, 0.1)', border: '1px solid rgba(139,92,246,0.3)', borderRadius: '12px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', fontWeight: 'bold' }}>
              <Loader2 className="animate-spin" size={16} style={{ color: 'hsl(var(--primary))' }} />
              Processando requisição...
            </div>
            <p style={{ fontSize: '11px', color: 'hsl(var(--text-muted))' }}>{statusMsg}</p>
          </div>
        )}

        {/* MESEP Quick Help Console */}
        <div className="glass-panel" style={{ padding: '20px', marginTop: 'auto', display: 'flex', flexDirection: 'column', gap: '12px' }}>
          <h4 style={{ fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>
            <HelpCircle size={14} style={{ color: 'hsl(var(--primary))' }} /> Regras de Qualidade MESEP
          </h4>
          <ul style={{ fontSize: '11px', color: 'hsl(var(--text-muted))', paddingLeft: '16px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <li><strong>Carga Horária:</strong> Verifique se a soma das horas das SAs bate com a carga total do curso (60h).</li>
            <li><strong>Critérios de Avaliação:</strong> Devem ser perguntas (Ex: "Desenvolveu o código sem erros de sintaxe?").</li>
            <li><strong>Critérios Críticos:</strong> Use o prefixo <code>[CRÍTICO]</code> na pergunta. Eles serão destacados em negrito no Word e PDF.</li>
            <li><strong>Estratégia:</strong> Defina problemas reais da profissão para engajar o aluno de forma ativa.</li>
          </ul>
        </div>
      </aside>
    </div>
  );
}
