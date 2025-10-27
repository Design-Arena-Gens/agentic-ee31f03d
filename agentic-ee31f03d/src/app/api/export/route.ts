import OpenApi from "@altravia/openapi";
import ExcelJS from "exceljs";
import { NextResponse } from "next/server";
import { z } from "zod";

const requestSchema = z.object({
  environment: z.enum(["test", "production"]),
  email: z.string().email(),
  apiKey: z.string().min(1, "API key obbligatoria"),
  token: z.string().min(1, "Token obbligatorio"),
  provincia: z.string().min(2).max(2),
  codiceAteco: z.string().min(1),
  denominazione: z.string().optional(),
  limit: z.number().min(1).max(500).optional(),
  skip: z.number().min(0).optional(),
});

type RequestPayload = z.infer<typeof requestSchema>;

type SearchFilters = Record<string, string | number>;

type ImpreseAdvanceResult = {
  denominazione?: string;
  piva?: string;
  cf?: string;
  indirizzo?: string;
  toponimo?: string;
  via?: string;
  civico?: string;
  cap?: string;
  comune?: string;
  provincia?: string;
  stato_attivita?: string;
  codice_destinatario?: string;
  id?: string;
  dettaglio?: {
    codice_ateco?: string;
    descrizione_ateco?: string;
    rea?: string;
    cciaa?: string;
    data_inizio_attivita?: string;
    pec?: string;
    cessata?: boolean;
  };
  gps?: {
    coordinates?: number[];
  };
};

export async function POST(request: Request) {
  let payload: RequestPayload;

  try {
    const data = await request.json();
    payload = requestSchema.parse(data);
  } catch (error) {
    const message =
      error instanceof z.ZodError
        ? error.issues[0]?.message ?? "Parametri non validi"
        : "Corpo della richiesta non valido";
    return NextResponse.json({ error: message }, { status: 400 });
  }

  const { environment, email, apiKey, token, provincia, codiceAteco, denominazione, limit, skip } =
    payload;

  try {
    const client = await OpenApi.init(environment, email, apiKey, token, true);

    if (!client.imprese) {
      return NextResponse.json(
        {
          error:
            "Il token fornito non include lo scope imprese.openapi.it. Genera un token con i permessi corretti.",
        },
        { status: 403 },
      );
    }

    const filters: SearchFilters = {};

    if (provincia) {
      filters.provincia = provincia.toUpperCase();
    }
    if (codiceAteco) {
      filters.codice_ateco = codiceAteco;
    }
    if (denominazione) {
      filters.denominazione = denominazione;
    }
    if (typeof limit === "number") {
      filters.limit = limit;
    }
    if (typeof skip === "number") {
      filters.skip = skip;
    }

    const rawResults = await client.imprese.search(filters as SearchFilters);

    if (!Array.isArray(rawResults)) {
      return NextResponse.json(
        { error: "La risposta del servizio non è nel formato atteso." },
        { status: 502 },
      );
    }

    const results = rawResults as ImpreseAdvanceResult[];

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "agentic-ee31f03d";
    workbook.created = new Date();

    const sheet = workbook.addWorksheet("Aziende");
    sheet.columns = [
      { header: "Denominazione", key: "denominazione", width: 40 },
      { header: "Partita IVA", key: "piva", width: 18 },
      { header: "Codice Fiscale", key: "cf", width: 20 },
      { header: "Indirizzo", key: "indirizzo", width: 40 },
      { header: "CAP", key: "cap", width: 10 },
      { header: "Comune", key: "comune", width: 22 },
      { header: "Provincia", key: "provincia", width: 12 },
      { header: "Codice ATECO", key: "codiceAteco", width: 15 },
      { header: "Descrizione ATECO", key: "descrizioneAteco", width: 32 },
      { header: "Stato attività", key: "statoAttivita", width: 18 },
      { header: "REA", key: "rea", width: 12 },
      { header: "CCIAA", key: "cciaa", width: 10 },
      { header: "Data inizio", key: "dataInizio", width: 16 },
      { header: "PEC", key: "pec", width: 30 },
      { header: "Codice destinatario", key: "codiceDestinatario", width: 18 },
      { header: "Cessata", key: "cessata", width: 10 },
      { header: "Latitudine", key: "lat", width: 14 },
      { header: "Longitudine", key: "lng", width: 14 },
      { header: "ID", key: "id", width: 20 },
    ];

    results.forEach((item) => {
      const detail = item.dettaglio ?? {};
      const coordinates = Array.isArray(item.gps?.coordinates)
        ? item.gps?.coordinates
        : [];
      const indirizzoCompiled =
        item.indirizzo ||
        [item.toponimo, item.via, item.civico].filter(Boolean).join(" ") ||
        "";

      sheet.addRow({
        denominazione: item.denominazione ?? "",
        piva: item.piva ?? "",
        cf: item.cf ?? "",
        indirizzo: indirizzoCompiled.trim(),
        cap: item.cap ?? "",
        comune: item.comune ?? "",
        provincia: item.provincia ?? "",
        codiceAteco: detail.codice_ateco ?? "",
        descrizioneAteco: detail.descrizione_ateco ?? "",
        statoAttivita: item.stato_attivita ?? "",
        rea: detail.rea ?? "",
        cciaa: detail.cciaa ?? "",
        dataInizio: detail.data_inizio_attivita ?? "",
        pec: detail.pec ?? "",
        codiceDestinatario: item.codice_destinatario ?? "",
        cessata: detail.cessata === true ? "SI" : detail.cessata === false ? "NO" : "",
        lat: typeof coordinates[1] === "number" ? coordinates[1] : "",
        lng: typeof coordinates[0] === "number" ? coordinates[0] : "",
        id: item.id ?? "",
      });
    });

    sheet.getRow(1).font = { bold: true };
    sheet.autoFilter = {
      from: "A1",
      to: "S1",
    };

    const buffer = await workbook.xlsx.writeBuffer();
    const binary = Buffer.from(buffer);

    const safeAteco = (codiceAteco || "ateco")
      .replace(/[^0-9a-z]+/gi, "_")
      .replace(/^_+|_+$/g, "")
      .toLowerCase();
    const filename = `imprese_${provincia.toUpperCase()}_${safeAteco || "codice"}_${new Date()
      .toISOString()
      .slice(0, 10)}.xlsx`;

    const headers = new Headers({
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${filename}"`,
      "Cache-Control": "no-store",
      "Content-Length": String(binary.length),
      "X-Result-Count": String(results.length),
    });

    return new Response(binary, { status: 200, headers });
  } catch (error: unknown) {
    const message = extractErrorMessage(error);
    const status = getStatusCode(error);
    return NextResponse.json({ error: message }, { status });
  }
}

function extractErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    // @ts-expect-error AxiosError compatibility
    const apiMessage = error.response?.data?.message ?? error.response?.data?.error;
    if (typeof apiMessage === "string" && apiMessage.trim().length > 0) {
      return apiMessage;
    }
    return error.message;
  }
  return "Errore inatteso durante l'elaborazione della richiesta.";
}

function getStatusCode(error: unknown): number {
  if (error instanceof Error) {
    // @ts-expect-error AxiosError compatibility
    const status = error.response?.status;
    if (typeof status === "number" && status >= 400 && status < 600) {
      return status;
    }
  }
  return 500;
}
