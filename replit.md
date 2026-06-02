# Radar de Inovação

Aplicativo web que converte planilhas Excel de projetos em apresentações PowerPoint do Radar de Inovação, com bolhas por categoria e slide de projetos destaque.

## Run & Operate

- `pnpm --filter @workspace/api-server run dev` — API server (port 8080)
- `pnpm --filter @workspace/radar-app run dev` — frontend (port 22353)
- `pnpm run typecheck` — full typecheck

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- API: Express 5 + multer (upload) + xlsx (parse) + jszip (PPTX)
- Frontend: React + Vite (sem DB — stateless)
- Build: esbuild (ESM bundle)

## Where things live

- `artifacts/api-server/src/routes/radar.ts` — endpoints parse + generate
- `artifacts/api-server/src/lib/pptxGenerator.ts` — lógica de geração PPTX com jszip
- `artifacts/radar-app/src/pages/Upload.tsx` — tela de upload
- `artifacts/radar-app/src/pages/Preview.tsx` — seleção de destaques + download
- `attached_assets/Radar_Inovacao_(3)_1780409449539.pptx` — template base do PowerPoint (vazio, sem bolhas)

## Architecture decisions

- **jszip + template**: O PPTX é gerado modificando o template original via jszip, preservando exatamente a identidade visual (radar, legendas, textos fixos, logo GEIN).
- **Stateless**: Nenhum banco de dados. Os dados parseados ficam no estado React do cliente; o endpoint /generate recebe tudo via JSON.
- **Direct fetch**: Sem codegen OpenAPI — upload de arquivo e download de blob binário não casam bem com React Query hooks gerados.
- **Posicionamento dinâmico**: Bolhas distribuídas angularmente com distância radial baseada no % direcionado. Algoritmo de repulsão anti-sobreposição (80 iterações).

## Product

1. Upload de planilha Excel com colunas CATEGORIA, TÍTULO, ETAPA
2. Parser filtra Cancelado/Identificado, classifica Direcionados vs. Em Andamento
3. Usuário vê estatísticas por categoria e seleciona projetos destaque
4. Download de PPTX com 2 slides: Slide 1 = radar com bolhas; Slide 2 = destaque com listas por categoria

## Gotchas

- O template PPTX está em `attached_assets/` — caminho hardcoded no pptxGenerator.ts
- No bundle esbuild, `__dirname` = `dist/` (3 níveis acima de `workspace/`)
- Etapas direcionadas: Concluído, Solução Experimentada
- Etapas excluídas: Cancelado, Identificado
