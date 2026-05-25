export interface CategoryData {
  name: string;
  displayName: string;
  total: number;
  directed: number;
  directedPct: number;
}

export interface Project {
  id: string;
  categoria: string;
  titulo: string;
  etapa: string;
  isDirected: boolean;
}

export interface ParseResult {
  projects: Project[];
  categories: CategoryData[];
  totalProjects: number;
  totalDirected: number;
  totalInProgress: number;
}

export interface HighlightProject {
  title: string;
  category: string;
}

export interface GeneratePayload {
  categories: CategoryData[];
  totalProjects: number;
  totalDirected: number;
  totalInProgress: number;
  highlights: HighlightProject[];
}
