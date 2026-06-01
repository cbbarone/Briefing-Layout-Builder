import { useState, useRef, useCallback } from "react";
import { parseExcel } from "../lib/api";
import type { ParseResult } from "../lib/types";

interface Props {
  onParsed: (result: ParseResult) => void;
}

export default function UploadPage({ onParsed }: Props) {
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const processFile = useCallback(
    async (file: File) => {
      if (
        !file.name.match(/\.(xlsx|xls)$/i)
      ) {
        setError("Envie um arquivo Excel (.xlsx ou .xls).");
        return;
      }
      setError(null);
      setLoading(true);
      try {
        const result = await parseExcel(file);
        onParsed(result);
      } catch (err: any) {
        setError(err.message || "Erro ao processar a planilha.");
      } finally {
        setLoading(false);
      }
    },
    [onParsed],
  );

  function onDrop(e: React.DragEvent) {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }

  function onFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (file) processFile(file);
    e.target.value = "";
  }

  return (
    <div className="page-upload">
      <header className="app-header">
        <div className="header-inner">
          <div className="header-logo">
            <span className="logo-dot" />
            <span className="logo-dot" style={{ background: "#006CB7" }} />
          </div>
          <h1 className="header-title">Radar de Inovação</h1>
        </div>
      </header>

      <main className="upload-main">
        <div className="upload-card">
          <div className="upload-card-header">
            <h2 className="upload-heading">Gerar Apresentação</h2>
            <p className="upload-subtext">
              Faça o upload da planilha de projetos para gerar o PowerPoint do
              Radar de Inovação automaticamente.
            </p>
          </div>

          <div
            className={`drop-zone${dragging ? " drop-zone--active" : ""}${loading ? " drop-zone--loading" : ""}`}
            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            onClick={() => !loading && inputRef.current?.click()}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={onFileChange}
              style={{ display: "none" }}
            />
            {loading ? (
              <div className="drop-loading">
                <div className="spinner" />
                <p>Processando planilha…</p>
              </div>
            ) : (
              <div className="drop-content">
                <div className="drop-icon">
                  <svg
                    width="48"
                    height="48"
                    viewBox="0 0 48 48"
                    fill="none"
                  >
                    <rect
                      x="8"
                      y="4"
                      width="32"
                      height="40"
                      rx="4"
                      fill="#E8F4FC"
                      stroke="#84CBE9"
                      strokeWidth="2"
                    />
                    <rect
                      x="14"
                      y="16"
                      width="20"
                      height="3"
                      rx="1.5"
                      fill="#006CB7"
                    />
                    <rect
                      x="14"
                      y="23"
                      width="20"
                      height="3"
                      rx="1.5"
                      fill="#006CB7"
                      opacity="0.6"
                    />
                    <rect
                      x="14"
                      y="30"
                      width="12"
                      height="3"
                      rx="1.5"
                      fill="#006CB7"
                      opacity="0.3"
                    />
                    <path
                      d="M30 2L38 10H30V2Z"
                      fill="#84CBE9"
                    />
                  </svg>
                </div>
                <p className="drop-text-main">
                  {dragging
                    ? "Solte o arquivo aqui"
                    : "Arraste sua planilha ou clique para selecionar"}
                </p>
                <p className="drop-text-sub">
                  Arquivos Excel (.xlsx, .xls) — colunas: CATEGORIAS, RESUMO, STATUS
                </p>
                <button
                  type="button"
                  className="btn btn-primary"
                  onClick={(e) => {
                    e.stopPropagation();
                    inputRef.current?.click();
                  }}
                >
                  Selecionar arquivo
                </button>
              </div>
            )}
          </div>

          {error && (
            <div className="error-box">
              <span className="error-icon">⚠</span>
              {error}
            </div>
          )}
        </div>

        <div className="upload-info">
          <h3>Como usar</h3>
          <ol className="steps-list">
            <li>
              <span className="step-num">1</span>
              <div>
                <strong>Faça o upload</strong> da planilha Excel com os projetos
                de inovação.
              </div>
            </li>
            <li>
              <span className="step-num">2</span>
              <div>
                <strong>Selecione os projetos destaque</strong> que aparecerão
                no segundo slide da apresentação.
              </div>
            </li>
            <li>
              <span className="step-num">3</span>
              <div>
                <strong>Gere e baixe</strong> o PowerPoint com o Radar de
                Inovação atualizado.
              </div>
            </li>
          </ol>
        </div>
      </main>
    </div>
  );
}
