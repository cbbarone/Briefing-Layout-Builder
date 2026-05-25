import { useState } from "react";
import { generatePptx, downloadBlob } from "../lib/api";
import type { ParseResult, HighlightProject } from "../lib/types";

interface Props {
  data: ParseResult;
  onReset: () => void;
}

export default function PreviewPage({ data, onReset }: Props) {
  const { projects, categories, totalProjects, totalDirected, totalInProgress } =
    data;

  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [generating, setGenerating] = useState(false);
  const [genError, setGenError] = useState<string | null>(null);
  const [expandedCat, setExpandedCat] = useState<string | null>(null);

  function toggleProject(id: string) {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleCategory(cat: string) {
    const catProjects = projects.filter((p) => p.categoria === cat);
    const allSelected = catProjects.every((p) => selected.has(p.id));
    setSelected((prev) => {
      const next = new Set(prev);
      if (allSelected) {
        catProjects.forEach((p) => next.delete(p.id));
      } else {
        catProjects.forEach((p) => next.add(p.id));
      }
      return next;
    });
  }

  async function handleGenerate() {
    setGenError(null);
    setGenerating(true);
    try {
      const highlights: HighlightProject[] = projects
        .filter((p) => selected.has(p.id))
        .map((p) => ({ title: p.titulo, category: p.categoria.toUpperCase() }));

      const blob = await generatePptx({
        categories,
        totalProjects,
        totalDirected,
        totalInProgress,
        highlights,
      });
      downloadBlob(blob, "Radar_Inovacao.pptx");
    } catch (err: any) {
      setGenError(err.message || "Erro ao gerar o PowerPoint.");
    } finally {
      setGenerating(false);
    }
  }

  function pct(n: number) {
    return `${Math.round(n * 100)}%`;
  }

  return (
    <div className="page-preview">
      <header className="app-header">
        <div className="header-inner">
          <button className="btn-back" onClick={onReset} title="Nova planilha">
            ← Nova planilha
          </button>
          <h1 className="header-title">Radar de Inovação</h1>
          <button
            className="btn btn-generate"
            onClick={handleGenerate}
            disabled={generating}
          >
            {generating ? (
              <>
                <span className="spinner spinner-sm" /> Gerando…
              </>
            ) : (
              <>
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: 6 }}>
                  <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
                  <polyline points="7 10 12 15 17 10"/>
                  <line x1="12" y1="15" x2="12" y2="3"/>
                </svg>
                Gerar PowerPoint
              </>
            )}
          </button>
        </div>
      </header>

      {genError && (
        <div className="error-box error-box--top">
          <span className="error-icon">⚠</span>
          {genError}
        </div>
      )}

      <main className="preview-main">
        <div className="stats-summary">
          <div className="stat-card stat-card--total">
            <span className="stat-num">{totalProjects}</span>
            <span className="stat-label">Total de Projetos</span>
          </div>
          <div className="stat-card stat-card--progress">
            <div className="stat-legend-dot" style={{ background: "#84CBE9" }} />
            <span className="stat-num">{totalInProgress}</span>
            <span className="stat-label">Em Andamento</span>
          </div>
          <div className="stat-card stat-card--directed">
            <div className="stat-legend-dot" style={{ background: "#006CB7" }} />
            <span className="stat-num">{totalDirected}</span>
            <span className="stat-label">Direcionados</span>
          </div>
          <div className="stat-card stat-card--cats">
            <span className="stat-num">{categories.length}</span>
            <span className="stat-label">Categorias</span>
          </div>
        </div>

        <div className="preview-grid">
          <section className="categories-section">
            <h2 className="section-title">
              Categorias
              <span className="section-badge">{categories.length}</span>
            </h2>
            <div className="categories-table-wrap">
              <table className="categories-table">
                <thead>
                  <tr>
                    <th>Categoria</th>
                    <th>Total</th>
                    <th>Direcionados</th>
                    <th>% Direcionado</th>
                    <th>Radar</th>
                  </tr>
                </thead>
                <tbody>
                  {categories.map((cat) => (
                    <tr key={cat.name}>
                      <td className="cat-name-cell">{cat.displayName}</td>
                      <td className="num-cell">{cat.total}</td>
                      <td className="num-cell">{cat.directed}</td>
                      <td className="num-cell">{pct(cat.directedPct)}</td>
                      <td>
                        <div className="radar-bar-wrap">
                          <div className="radar-bar">
                            <div
                              className="radar-bar-fill"
                              style={{ width: pct(cat.directedPct) }}
                            />
                          </div>
                          <span className="radar-bar-label">
                            {cat.directedPct >= 0.66
                              ? "🔵 Centro"
                              : cat.directedPct >= 0.33
                                ? "🟦 Médio"
                                : "🩵 Externo"}
                          </span>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <section className="highlights-section">
            <h2 className="section-title">
              Projetos Destaque
              <span className="section-badge">{selected.size} selecionados</span>
            </h2>
            <p className="highlights-hint">
              Selecione os projetos que aparecerão no 2º slide da apresentação.
            </p>
            <div className="projects-list">
              {categories.map((cat) => {
                const catProjects = projects.filter(
                  (p) => p.categoria === cat.name,
                );
                const isExpanded = expandedCat === cat.name;
                const allSelected = catProjects.every((p) =>
                  selected.has(p.id),
                );
                const someSelected = catProjects.some((p) =>
                  selected.has(p.id),
                );

                return (
                  <div key={cat.name} className="cat-group">
                    <div
                      className={`cat-group-header${isExpanded ? " cat-group-header--open" : ""}`}
                    >
                      <label className="cat-check-label">
                        <input
                          type="checkbox"
                          className="checkbox"
                          checked={allSelected}
                          ref={(el) => {
                            if (el)
                              el.indeterminate = someSelected && !allSelected;
                          }}
                          onChange={() => toggleCategory(cat.name)}
                        />
                        <span className="cat-group-name">{cat.displayName}</span>
                        <span className="cat-group-count">
                          {catProjects.filter((p) => selected.has(p.id)).length ||
                            ""}
                          /{catProjects.length}
                        </span>
                      </label>
                      <button
                        className="expand-btn"
                        onClick={() =>
                          setExpandedCat(isExpanded ? null : cat.name)
                        }
                        aria-label={isExpanded ? "Recolher" : "Expandir"}
                      >
                        {isExpanded ? "▲" : "▼"}
                      </button>
                    </div>

                    {isExpanded && (
                      <div className="project-items">
                        {catProjects.map((p) => (
                          <label key={p.id} className="project-item">
                            <input
                              type="checkbox"
                              className="checkbox"
                              checked={selected.has(p.id)}
                              onChange={() => toggleProject(p.id)}
                            />
                            <div className="project-item-info">
                              <span className="project-title">{p.titulo}</span>
                              <span
                                className={`project-badge${p.isDirected ? " project-badge--directed" : ""}`}
                              >
                                {p.isDirected ? "Direcionado" : "Em Andamento"}
                              </span>
                            </div>
                          </label>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </section>
        </div>

        <div className="generate-bottom">
          <div className="generate-summary">
            <strong>{totalProjects}</strong> projetos em{" "}
            <strong>{categories.length}</strong> categorias ·{" "}
            <strong>{selected.size}</strong> destaque
            {selected.size !== 1 ? "s" : ""} selecionado
            {selected.size !== 1 ? "s" : ""}
          </div>
          <button
            className="btn btn-generate btn-generate--large"
            onClick={handleGenerate}
            disabled={generating}
          >
            {generating ? (
              <>
                <span className="spinner spinner-sm" /> Gerando…
              </>
            ) : (
              "⬇ Gerar e Baixar PowerPoint"
            )}
          </button>
        </div>
      </main>
    </div>
  );
}
