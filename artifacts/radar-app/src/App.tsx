import { useState } from "react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import UploadPage from "./pages/Upload";
import PreviewPage from "./pages/Preview";
import type { ParseResult } from "./lib/types";

const queryClient = new QueryClient();

type Step = "upload" | "preview";

function RadarApp() {
  const [step, setStep] = useState<Step>("upload");
  const [parsed, setParsed] = useState<ParseResult | null>(null);

  function handleParsed(result: ParseResult) {
    setParsed(result);
    setStep("preview");
  }

  function handleReset() {
    setParsed(null);
    setStep("upload");
  }

  return (
    <div className="app-root">
      {step === "upload" && <UploadPage onParsed={handleParsed} />}
      {step === "preview" && parsed && (
        <PreviewPage data={parsed} onReset={handleReset} />
      )}
    </div>
  );
}

export default function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <RadarApp />
    </QueryClientProvider>
  );
}
