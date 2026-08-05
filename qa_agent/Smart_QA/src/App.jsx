import { useMemo, useRef, useState } from 'react'
import './App.css'

const cards = [
  {
    title: 'Generate from Excel',
    subtitle: 'Generate prompt-pack test cases from requirement sheets.',
    icon: 'XL',
    variant: 'primary',
  },
  {
    title: 'Generate from DOCX',
    subtitle: 'Extract requirements from documents and generate cases.',
    icon: 'DOC',
    variant: 'secondary',
  },
  {
    title: 'Generate from Image',
    subtitle: 'Infer flows from screenshots and generate QA coverage.',
    icon: 'IMG',
    variant: 'secondary',
  },
  {
    title: 'Review Results',
    subtitle: 'Run coverage checks on generated test case outputs.',
    icon: 'QA',
    variant: 'outline',
  },
]

const apiBases = ['', 'http://127.0.0.1:8000']

function App() {
  const [selected, setSelected] = useState('Generate from Excel')
  const [selectedFile, setSelectedFile] = useState(null)
  const [status, setStatus] = useState('Ready')
  const [message, setMessage] = useState('Choose a file to begin the workflow.')
  const [summary, setSummary] = useState('')
  const [downloadInfo, setDownloadInfo] = useState(null)
  const [isRunning, setIsRunning] = useState(false)
  const [isDownloading, setIsDownloading] = useState(false)
  const [isDragging, setIsDragging] = useState(false)
  const [llmEndpoint, setLlmEndpoint] = useState('')
  const [apiKey, setApiKey] = useState('')
  const [model, setModel] = useState('')
  const [showLlmSettings, setShowLlmSettings] = useState(false)
  const [llmConnectionStatus, setLlmConnectionStatus] = useState(null)
  const [isTestingLlm, setIsTestingLlm] = useState(false)
  const fileInputRef = useRef(null)

  const activeCard = useMemo(() => cards.find((card) => card.title === selected) ?? cards[0], [selected])

  const acceptValue = useMemo(() => {
    switch (activeCard.title) {
      case 'Generate from Excel':
      case 'Review Results':
        return '.xlsx,.xls'
      case 'Generate from DOCX':
        return '.docx'
      case 'Generate from Image':
        return 'image/*'
      default:
        return '*/*'
    }
  }, [activeCard.title])

  const usesLlmGeneration = activeCard.title !== 'Review Results'

  const endpoint = useMemo(() => {
    switch (activeCard.title) {
      case 'Generate from Excel':
        return '/api/generate-from-excel'
      case 'Generate from DOCX':
        return '/api/generate-from-docx'
      case 'Generate from Image':
        return '/api/generate-from-image'
      case 'Review Results':
        return '/api/review-results'
      default:
        return '/api/health'
    }
  }, [activeCard.title])

  const handleFileChange = (event) => {
    const file = event.target.files?.[0] ?? null
    setSelectedFile(file)
    setStatus(file ? 'Ready' : status)
    setMessage(file ? `Selected ${file.name}` : 'Choose a file to begin the workflow.')
    setDownloadInfo(null)
  }

  const runWorkflowRequest = async (formData) => {
    let lastError = null

    for (const base of apiBases) {
      try {
        const response = await fetch(`${base}${endpoint}`, {
          method: 'POST',
          body: formData,
        })
        const payload = await response.json().catch(() => ({}))

        if (!response.ok) {
          throw new Error(payload.detail ?? payload.message ?? `Request failed with status ${response.status}.`)
        }

        return payload
      } catch (error) {
        lastError = error
        if (base) {
          break
        }
      }
    }

    throw lastError ?? new Error('The workflow failed.')
  }

  const handleTestLlmConnection = async () => {
    if (!llmEndpoint.trim() || !apiKey.trim()) {
      setLlmConnectionStatus({
        type: 'error',
        text: 'Enter the LLM endpoint and API key before testing.',
      })
      return
    }

    setIsTestingLlm(true)
    setLlmConnectionStatus({ type: 'running', text: 'Testing LLM connection...' })

    const formData = new FormData()
    formData.append('llm_endpoint', llmEndpoint.trim())
    formData.append('api_key', apiKey.trim())
    formData.append('model', model.trim())

    let lastError = null
    for (const base of apiBases) {
      try {
        const response = await fetch(`${base}/api/test-llm-connection`, {
          method: 'POST',
          body: formData,
        })
        const payload = await response.json().catch(() => ({}))
        if (!response.ok) {
          throw new Error(payload.detail ?? payload.message ?? `Connection failed with status ${response.status}.`)
        }
        setLlmConnectionStatus({
          type: 'success',
          text: payload.message ?? 'LLM connection successful.',
        })
        setMessage(payload.message ?? 'LLM connection successful.')
        setIsTestingLlm(false)
        return
      } catch (error) {
        lastError = error
        if (base) {
          break
        }
      }
    }

    setLlmConnectionStatus({
      type: 'error',
      text: lastError?.message ?? 'LLM connection failed.',
    })
    setIsTestingLlm(false)
  }

  const handleStart = async () => {
    if (!selectedFile) {
      setStatus('Needs file')
      setMessage('Please choose a file before starting the workflow.')
      setSummary('')
      setDownloadInfo(null)
      return
    }

    setIsRunning(true)
    setStatus('Running')
    setMessage(`Processing ${selectedFile.name} with ${activeCard.title}...`)
    setSummary('')
    setDownloadInfo(null)

    const formData = new FormData()
    formData.append('file', selectedFile)
    if (usesLlmGeneration) {
      formData.append('llm_endpoint', llmEndpoint.trim())
      formData.append('api_key', apiKey.trim())
      formData.append('model', model.trim())
    }

    try {
      const payload = await runWorkflowRequest(formData)

      setStatus('Success')
      setMessage(payload.message ?? 'Workflow completed successfully.')
      setSummary(payload.summary ?? '')
      setDownloadInfo(
        payload.download_file
          ? {
              file: payload.download_file,
              name: payload.download_name ?? payload.download_file,
              label: payload.download_label ?? 'Download generated test cases',
            }
          : null,
      )
    } catch (error) {
      setStatus('Error')
      setMessage(error.message ?? 'An unexpected error occurred during processing.')
      setSummary('')
      setDownloadInfo(null)
    } finally {
      setIsRunning(false)
    }
  }

  const fetchDownloadBlob = async (filename) => {
    let lastError = null

    for (const base of apiBases) {
      try {
        const response = await fetch(`${base}/api/download/${encodeURIComponent(filename)}`)
        if (!response.ok) {
          const payload = await response.json().catch(() => ({}))
          throw new Error(payload.detail ?? `Download failed with status ${response.status}.`)
        }
        return await response.blob()
      } catch (error) {
        lastError = error
        if (base) {
          break
        }
      }
    }

    throw lastError ?? new Error('Download failed.')
  }

  const handleDownload = async () => {
    if (!downloadInfo) {
      return
    }

    setIsDownloading(true)
    try {
      const blob = await fetchDownloadBlob(downloadInfo.file)

      if ('showSaveFilePicker' in window) {
        const fileHandle = await window.showSaveFilePicker({
          suggestedName: downloadInfo.name,
          types: [
            {
              description: 'Excel workbook',
              accept: {
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
              },
            },
          ],
        })
        const writable = await fileHandle.createWritable()
        await writable.write(blob)
        await writable.close()
        setMessage(`Saved ${downloadInfo.name}`)
      } else {
        const url = URL.createObjectURL(blob)
        const anchor = document.createElement('a')
        anchor.href = url
        anchor.download = downloadInfo.name
        document.body.append(anchor)
        anchor.click()
        anchor.remove()
        URL.revokeObjectURL(url)
        setMessage('Download started. Your browser may use its default download folder.')
      }
    } catch (error) {
      if (error.name !== 'AbortError') {
        setStatus('Error')
        setMessage(error.message ?? 'Could not download the generated test cases.')
      }
    } finally {
      setIsDownloading(false)
    }
  }

  const handlePreview = () => {
    fileInputRef.current?.click()
  }

  const handleDrop = (event) => {
    event.preventDefault()
    setIsDragging(false)
    const file = event.dataTransfer.files?.[0] ?? null
    if (!file) {
      return
    }
    setSelectedFile(file)
    setStatus('Ready')
    setMessage(`Selected ${file.name}`)
    setDownloadInfo(null)
  }

  return (
    <main className="app-shell">
      <section className="app-frame shadow-lg">
        <div className="app-header d-flex align-items-center gap-3">
          <div className="brand-badge d-flex align-items-center justify-content-center text-white fw-bold fs-3">
            Q
          </div>
          <div>
            <div className="app-title fw-semibold">QA Agent Studio</div>
            <div className="app-tagline text-muted">Prompt-pack + LLM Foundry test case generation</div>
          </div>
        </div>

        <div className="app-body">
          <div className="panel workflow-panel">
            <div className="d-grid gap-3">
              {cards.map((card) => (
                <button
                  key={card.title}
                  type="button"
                  className={`action-card btn d-flex align-items-center justify-content-between rounded-3 ${
                    selected === card.title ? 'selected' : ''
                  } ${card.variant}`}
                  onClick={() => {
                    setSelected(card.title)
                    setSelectedFile(null)
                    setStatus('Ready')
                    setMessage('Choose a file to begin the workflow.')
                    setSummary('')
                    setDownloadInfo(null)
                    if (fileInputRef.current) {
                      fileInputRef.current.value = ''
                    }
                  }}
                >
                  <span className="d-flex align-items-center gap-3">
                    <span className="action-icon rounded-3">{card.icon}</span>
                    <span className="text-start">
                      <span className="d-block fw-semibold text-dark">{card.title}</span>
                      <small className="text-muted d-block">{card.subtitle}</small>
                    </span>
                  </span>
                  <span className="chevron">&gt;</span>
                </button>
              ))}
            </div>
          </div>

          <div className="workspace-card">
            <div className="d-flex align-items-center justify-content-between mb-3">
              <div>
                <span className="workspace-label text-uppercase small fw-semibold">Current workflow</span>
                <h2 className="workspace-title h5 mb-0 mt-1">{activeCard.title}</h2>
              </div>
              <span className={`status-pill rounded-pill px-3 py-1 small fw-semibold ${status.toLowerCase()}`}>{status}</span>
            </div>
            <p className="workspace-desc text-muted mb-3">{activeCard.subtitle}</p>
            <div className="pipeline-strip">
              <span>Prompt Pack Rules</span>
              <span>LLM Foundry</span>
              <span>Excel Output</span>
              <span>Reviewer</span>
            </div>

            {usesLlmGeneration ? (
              <div className="llm-settings">
                <button
                  type="button"
                  className="llm-settings-toggle"
                  onClick={() => setShowLlmSettings((current) => !current)}
                >
                  {showLlmSettings ? 'Hide LLM settings' : 'Configure LLM Foundry'}
                </button>

                {showLlmSettings ? (
                  <div className="llm-settings-grid">
                    <label>
                      <span>LLM Endpoint</span>
                      <input
                        type="url"
                        value={llmEndpoint}
                        onChange={(event) => {
                          setLlmEndpoint(event.target.value)
                          setLlmConnectionStatus(null)
                        }}
                        placeholder="https://llm-foundry-host/v1/chat/completions"
                      />
                    </label>
                    <label>
                      <span>API Key</span>
                      <input
                        type="password"
                        value={apiKey}
                        onChange={(event) => {
                          setApiKey(event.target.value)
                          setLlmConnectionStatus(null)
                        }}
                        placeholder="Paste API key for this session"
                        autoComplete="off"
                      />
                    </label>
                    <label>
                      <span>Model</span>
                      <input
                        type="text"
                        value={model}
                        onChange={(event) => {
                          setModel(event.target.value)
                          setLlmConnectionStatus(null)
                        }}
                        placeholder="gpt-4o-mini"
                      />
                    </label>
                    <button
                      type="button"
                      className="llm-test-btn"
                      onClick={handleTestLlmConnection}
                      disabled={isTestingLlm}
                    >
                      {isTestingLlm ? 'Testing...' : 'Test LLM connection'}
                    </button>
                    {llmConnectionStatus ? (
                      <div className={`llm-connection-message ${llmConnectionStatus.type}`}>
                        {llmConnectionStatus.text}
                      </div>
                    ) : null}
                  </div>
                ) : null}
              </div>
            ) : null}

            <div
              className={`file-dropzone ${isDragging ? 'dragging' : ''} ${selectedFile ? 'has-file' : ''}`}
              onDragOver={(event) => {
                event.preventDefault()
                setIsDragging(true)
              }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={handleDrop}
            >
              <div className="file-drop-icon">{activeCard.icon}</div>
              <div className="file-drop-content">
                <div className="file-drop-title">{selectedFile ? selectedFile.name : 'Drop your file here'}</div>
                <div className="file-drop-subtitle">
                  {selectedFile ? 'Ready to run this workflow.' : `Accepted format: ${acceptValue}`}
                </div>
              </div>
              <button type="button" className="choose-file-btn" onClick={handlePreview}>
                {selectedFile ? 'Change file' : 'Choose file'}
              </button>
            </div>

            <div className="file-picker d-flex flex-column gap-2 mb-3">
              <label
                htmlFor="file-upload"
                className="visually-hidden"
              >
                {selectedFile ? 'Change file' : 'Choose file'}
              </label>
              <input
                id="file-upload"
                ref={fileInputRef}
                type="file"
                accept={acceptValue}
                onChange={handleFileChange}
                style={{ position: 'absolute', opacity: 0, width: 0, height: 0 }}
              />
            </div>

            <div className={`workflow-message ${status.toLowerCase().replace(/\s+/g, '-')}`}>{message}</div>

            {summary ? (
              <div className="summary-panel rounded-3 px-3 py-3 mb-4">
                <div className="summary-title small fw-semibold mb-2">Review Summary</div>
                <pre className="summary-text mb-0">{summary}</pre>
              </div>
            ) : null}

            <div className="action-row">
              <button className="start-btn" onClick={handleStart} disabled={isRunning}>
                {isRunning ? 'Running...' : 'Start'}
              </button>
              <button className="secondary-btn" onClick={handlePreview}>Browse</button>
              {downloadInfo ? (
                <button className="download-btn" onClick={handleDownload} disabled={isDownloading}>
                  {isDownloading ? 'Preparing...' : downloadInfo.label}
                </button>
              ) : null}
            </div>
          </div>
        </div>
      </section>
    </main>
  )
}

export default App
