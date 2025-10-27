"use client";

import { ChangeEvent, FormEvent, useMemo, useState } from "react";

type Environment = "test" | "production";
type Status =
  | { type: "idle" }
  | { type: "loading" }
  | { type: "error"; message: string }
  | { type: "success"; message: string };

const MAX_LIMIT = 500;

export default function Home() {
  const [formData, setFormData] = useState({
    environment: "test" as Environment,
    email: "",
    apiKey: "",
    token: "",
    provincia: "VR",
    codiceAteco: "10.71",
    limit: "100",
    skip: "",
    denominazione: "",
  });
  const [status, setStatus] = useState<Status>({ type: "idle" });
  const [resultCount, setResultCount] = useState<number | null>(null);
  const [lastFileName, setLastFileName] = useState<string | null>(null);

  const limitHint = useMemo(() => {
    const raw = Number(formData.limit || "0");
    if (!raw || Number.isNaN(raw)) return "";
    if (raw > MAX_LIMIT) {
      return `Il limite massimo consentito è ${MAX_LIMIT}. Verrà applicato automaticamente.`;
    }
    return "";
  }, [formData.limit]);

  const handleChange =
    (field: keyof typeof formData) =>
    (event: ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
      setFormData((prev) => ({
        ...prev,
        [field]: event.target.value,
      }));
    };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setStatus({ type: "loading" });
    setResultCount(null);
    setLastFileName(null);

    try {
      const rawLimit = formData.limit ? Number(formData.limit) : undefined;
      const limit = rawLimit
        ? Math.min(Math.max(Math.floor(rawLimit), 1), MAX_LIMIT)
        : undefined;
      const skip = formData.skip
        ? Math.max(Math.floor(Number(formData.skip)), 0)
        : undefined;

      const payload = {
        environment: formData.environment,
        email: formData.email.trim(),
        apiKey: formData.apiKey.trim(),
        token: formData.token.trim(),
        provincia: formData.provincia.trim().toUpperCase(),
        codiceAteco: formData.codiceAteco.trim(),
        denominazione: formData.denominazione.trim() || undefined,
        limit,
        skip,
      };

      const response = await fetch("/api/export", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        let message = "Impossibile completare l'estrazione.";
        try {
          const data = await response.json();
          if (data?.error) {
            message = data.error;
          }
        } catch {
          // ignored
        }
        throw new Error(message);
      }

      const blob = await response.blob();
      const downloadUrl = window.URL.createObjectURL(blob);

      const safeAteco = (formData.codiceAteco || "ateco")
        .replace(/[^0-9a-z]+/gi, "_")
        .replace(/^_+|_+$/g, "")
        .toLowerCase();
      const filename = `aziende_${payload.provincia}_${safeAteco || "codice"}_${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`;

      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(downloadUrl);

      const count = Number(response.headers.get("x-result-count") || "0");
      setResultCount(Number.isNaN(count) ? null : count);
      setLastFileName(filename);
      setStatus({
        type: "success",
        message: "File Excel generato correttamente.",
      });
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : "Si è verificato un errore inatteso.";
      setStatus({ type: "error", message });
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 py-12">
      <main className="mx-auto flex w-full max-w-5xl flex-col gap-8 rounded-xl border border-slate-200 bg-white p-8 shadow-sm">
        <header className="flex flex-col gap-2">
          <span className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-500">
            Openapi · Imprese
          </span>
          <h1 className="text-3xl font-semibold text-slate-900">
            Estrai aziende per codice ATECO e provincia
          </h1>
          <p className="max-w-3xl text-sm text-slate-600">
            Inserisci le credenziali del tuo account Openapi e imposta i filtri di ricerca. Il servizio{" "}
            invia una chiamata all&apos;endpoint{" "}
            <code className="rounded bg-slate-100 px-1 py-0.5 text-xs">/advance</code>{" "}
            di{" "}
            <code className="rounded bg-slate-100 px-1 py-0.5 text-xs">imprese.openapi.it</code>{" "}
            e genera un foglio Excel scaricabile con i risultati.
          </p>
        </header>

        <section className="rounded-lg border border-slate-200 bg-slate-50 px-6 py-4 text-sm text-slate-600">
          <ul className="list-disc space-y-2 pl-5">
            <li>
              I dati personali inseriti vengono utilizzati esclusivamente per questa richiesta e non vengono salvati.
            </li>
            <li>
              In modalità sandbox l&apos;endpoint è{" "}
              <code className="rounded bg-slate-100 px-1 py-0.5 text-xs">https://test.imprese.openapi.it</code>{" "}
              con le aziende dimostrative fornite da Openapi.
            </li>
            <li>
              Il token deve includere lo scope{" "}
              <code className="rounded bg-slate-100 px-1 py-0.5 text-xs">imprese.openapi.it</code>.
            </li>
          </ul>
        </section>

        <form className="grid grid-cols-1 gap-6 lg:grid-cols-2" onSubmit={handleSubmit}>
          <div className="flex flex-col gap-2">
            <label className="text-sm font-medium text-slate-800">Ambiente</label>
            <div className="flex gap-3 rounded-lg border border-slate-200 bg-white p-2 text-sm">
              <label className="flex w-full cursor-pointer select-none items-center gap-2 rounded-md px-3 py-2 hover:bg-slate-100">
                <input
                  type="radio"
                  name="environment"
                  value="test"
                  checked={formData.environment === "test"}
                  onChange={handleChange("environment")}
                  className="h-4 w-4 accent-slate-900"
                />
                Sandbox (test.imprese.openapi.it)
              </label>
              <label className="flex w-full cursor-pointer select-none items-center gap-2 rounded-md px-3 py-2 hover:bg-slate-100">
                <input
                  type="radio"
                  name="environment"
                  value="production"
                  checked={formData.environment === "production"}
                  onChange={handleChange("environment")}
                  className="h-4 w-4 accent-slate-900"
                />
                Produzione (imprese.openapi.it)
              </label>
            </div>
          </div>

          <div className="flex flex-col gap-2">
            <label className="text-sm font-medium text-slate-800">Email account Openapi</label>
            <input
              required
              type="email"
              placeholder="nome@azienda.it"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200"
              value={formData.email}
              onChange={handleChange("email")}
            />
          </div>

          <div className="flex flex-col gap-2">
            <label className="text-sm font-medium text-slate-800">API Key</label>
            <input
              required
              type="password"
              placeholder="API key"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200"
              value={formData.apiKey}
              onChange={handleChange("apiKey")}
            />
          </div>

          <div className="flex flex-col gap-2">
            <label className="text-sm font-medium text-slate-800">Token OAuth</label>
            <input
              required
              type="password"
              placeholder="Token generato con scope imprese.openapi.it"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200"
              value={formData.token}
              onChange={handleChange("token")}
            />
          </div>

          <div className="flex flex-col gap-2">
            <label className="text-sm font-medium text-slate-800">Provincia (es. VR)</label>
            <input
              required
              maxLength={2}
              placeholder="Sigla provincia"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm uppercase shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200"
              value={formData.provincia}
              onChange={handleChange("provincia")}
            />
          </div>

          <div className="flex flex-col gap-2">
            <label className="text-sm font-medium text-slate-800">Codice ATECO</label>
            <input
              required
              placeholder="Es. 10.71 o 1071"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200"
              value={formData.codiceAteco}
              onChange={handleChange("codiceAteco")}
            />
          </div>

          <div className="flex flex-col gap-2">
            <label className="text-sm font-medium text-slate-800">Denominazione (facoltativo)</label>
            <input
              placeholder="Filtra per denominazione"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200"
              value={formData.denominazione}
              onChange={handleChange("denominazione")}
            />
          </div>

          <div className="flex flex-col gap-2">
            <label className="text-sm font-medium text-slate-800">Limite risultati (max 500)</label>
            <input
              required
              type="number"
              min={1}
              max={MAX_LIMIT}
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200"
              value={formData.limit}
              onChange={handleChange("limit")}
            />
            {limitHint && <p className="text-xs text-slate-500">{limitHint}</p>}
          </div>

          <div className="flex flex-col gap-2">
            <label className="text-sm font-medium text-slate-800">Skip (impaginazione)</label>
            <input
              type="number"
              min={0}
              placeholder="Numero di risultati da saltare"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200"
              value={formData.skip}
              onChange={handleChange("skip")}
            />
          </div>

          <div className="lg:col-span-2">
            <button
              type="submit"
              disabled={status.type === "loading"}
              className="inline-flex w-full items-center justify-center rounded-lg bg-slate-900 px-6 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 focus:outline-none focus:ring-2 focus:ring-slate-300 disabled:cursor-not-allowed disabled:bg-slate-400 lg:w-auto"
            >
              {status.type === "loading" ? "Generazione in corso..." : "Genera file Excel"}
            </button>
          </div>
        </form>

        <section className="min-h-[2.5rem] rounded-lg border border-slate-200 bg-slate-50 px-5 py-3 text-sm text-slate-700">
          {status.type === "idle" && (
            <p>
              Compila il modulo con le tue credenziali e premi <strong>Genera file Excel</strong>.
            </p>
          )}
          {status.type === "loading" && (
            <p className="animate-pulse font-medium text-slate-900">
              Connessione al servizio Openapi in corso...
            </p>
          )}
          {status.type === "success" && (
            <div className="space-y-1 text-slate-700">
              <p className="font-semibold text-emerald-600">{status.message}</p>
              {typeof resultCount === "number" && (
                <p>
                  Aziende incluse nel file:{" "}
                  <span className="font-medium text-slate-900">{resultCount}</span>
                </p>
              )}
              {lastFileName && (
                <p>
                  File scaricato:{" "}
                  <span className="font-mono text-slate-900">{lastFileName}</span>
                </p>
              )}
            </div>
          )}
          {status.type === "error" && (
            <p className="font-semibold text-red-600">{status.message}</p>
          )}
        </section>
      </main>
    </div>
  );
}
